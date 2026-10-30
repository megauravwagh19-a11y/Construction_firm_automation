"""
TRIKON - Construction Engineering Studio
Windows desktop app (Python + Tkinter). Tabs: Letterheads, BOQ Generator,
Contract Generator, Material & Specs, Project Scope.

Install once:   pip install reportlab openpyxl
Run:            python trikon_studio.py
Optional logo:  put "logo.png" (transparent background) next to this file.
"""
import os
import datetime
import csv
import json
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from xml.sax.saxutils import escape as xml_escape
from reportlab.pdfgen import canvas as canvas_mod
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# ----------------------------------------------------------------- firm info
FIRM = {
    "name": "TRIKON",
    "sub": "CONSTRUCTION & DESIGNS",
    "tag": "DESIGN \u2022 DEVELOP \u2022 DELIVER",
    "owner": "Er. Gaurav S. Wagh",
    "owner_full": "Er. Gaurav Somnath Wagh",
    "role": "OWNER & CIVIL ENGINEER",
    "quals": "B.Tech Civil \u2022 Master Cert. in Building Design",
    "phone": "+91 9699354459",
    "email": "gaurav.adc09@gmail.com",
    "city": "Pandharpur - 413304",
    "address": "Pl. No. 58, Chhatrapati Sambhaji Maharaj Nagar, Near New Kane Hospital, "
               "Thakare Chowk, Pandharpur, Dist. Solapur - 413304",
}
NAVY, GOLD, MAROON = "#161b26", "#c9a227", "#7a1f2b"
HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "logo.png")

# --------------------------------------------------------------- PDF helpers
FONT, FONT_B, RS = "Helvetica", "Helvetica-Bold", "Rs."
for reg, bold in [(r"C:\Windows\Fonts\segoeui.ttf", r"C:\Windows\Fonts\segoeuib.ttf")]:
    try:
        pdfmetrics.registerFont(TTFont("Seg", reg))
        pdfmetrics.registerFont(TTFont("Seg-B", bold))
        FONT, FONT_B, RS = "Seg", "Seg-B", "\u20b9"
    except Exception:
        pass

C_GOLD, C_NAVY, C_MAROON = colors.HexColor(GOLD), colors.HexColor(NAVY), colors.HexColor(MAROON)


def money(x):
    return f"{x:,.2f}"


def today():
    return datetime.date.today().strftime("%d/%m/%Y")


def safe_name(text, limit=60):
    t = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "-", str(text)).strip(" .-")
    t = re.sub(r"\s+", "_", t)
    return (t or "Unnamed")[:limit]


class Store:
    """All data stays on this PC.

    Projects folder layout:
      Trikon_Projects/
        1_Letterheads/<Client>_<Contact>/<Client>_<Ref>.pdf
        2_BOQ/...  3_Contracts/...  4_Material_Specs/...  5_Project_Scope/...
        _data/trikon_data.json   (ref counters + saved clients)
        _data/Document_Register.csv   (log of every document made)
    """
    KINDS = {
        "letter": ("1_Letterheads", "TCD/GEN/", "Letterhead"),
        "boq": ("2_BOQ", "BOQ-", "BOQ"),
        "contract": ("3_Contracts", "TCD/AGR/", "Contract"),
        "specs": ("4_Material_Specs", "MAT-SPEC-", "Material Specs"),
        "scope": ("5_Project_Scope", "SCP-", "Project Scope"),
    }

    def __init__(self):
        self.root = "D:\\Trikon_Projects" if os.path.exists("D:\\") else os.path.join(
            os.path.expanduser("~"), "Documents", "Trikon_Projects")
        self.data_dir = os.path.join(self.root, "_data")
        self.data_file = os.path.join(self.data_dir, "trikon_data.json")
        self.register = os.path.join(self.data_dir, "Document_Register.csv")
        self.data = {"counters": {}, "clients": {}}
        try:
            os.makedirs(self.data_dir, exist_ok=True)
            for sub, _, _ in self.KINDS.values():
                os.makedirs(os.path.join(self.root, sub), exist_ok=True)
            if os.path.exists(self.data_file):
                with open(self.data_file, "r", encoding="utf-8") as f:
                    self.data.update(json.load(f))
        except Exception:
            pass

    def save(self):
        try:
            tmp = self.data_file + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.data_file)
        except Exception as e:
            messagebox.showwarning("Data not saved", f"Could not save app data:\n{e}")

    def next_ref(self, kind):
        _, prefix, _ = self.KINDS[kind]
        d = datetime.date.today().strftime("%Y%m%d")
        n = self.data["counters"].get(f"{kind}:{d}", 0) + 1
        return f"{prefix}{d}-{n:02d}"

    def prepare(self, kind, client, contact, ref, ext=".pdf"):
        """Validate, create folders and return the full file path (or None)."""
        if not str(client).strip():
            messagebox.showwarning("Client name needed", "Please enter the client name first.")
            return None
        if not str(ref).strip():
            messagebox.showwarning("Ref. number needed", "Reference number is empty.")
            return None
        digits = re.sub(r"[^0-9+]", "", str(contact))
        if not digits:
            if not messagebox.askyesno("Contact number empty",
                                       "Contact number is empty.\nThe folder will be named without it.\n\nContinue?"):
                return None
            digits = "NoContact"
        sub = self.KINDS[kind][0]
        folder = os.path.join(self.root, sub, f"{safe_name(client)}_{digits}")
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Folder error", f"Could not create folder:\n{folder}\n\n{e}")
            return None
        base = f"{safe_name(client)}_{safe_name(ref, 40)}"
        path = os.path.join(folder, base + ext)
        n = 2
        while os.path.exists(path):
            path = os.path.join(folder, f"{base}_v{n}{ext}")
            n += 1
        return path

    def add_client(self, client, contact, site):
        key = f"{str(client).strip()} - {str(contact).strip()}" if str(contact).strip() else str(client).strip()
        self.data["clients"][key] = {"client": str(client).strip(), "contact": str(contact).strip(),
                                     "site": str(site).strip()}

    def client_keys(self):
        return sorted(self.data["clients"].keys(), key=str.lower)

    def log(self, kind, client, contact, ref, path):
        try:
            new = not os.path.exists(self.register)
            with open(self.register, "a", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                if new:
                    w.writerow(["Date", "Time", "Document", "Client", "Contact", "Ref No.", "File"])
                now = datetime.datetime.now()
                w.writerow([now.strftime("%d/%m/%Y"), now.strftime("%H:%M:%S"), self.KINDS[kind][2],
                            client, contact, ref, path])
        except Exception:
            pass

    def record(self, kind, client, contact, ref, site, path):
        _, prefix, _ = self.KINDS[kind]
        d = datetime.date.today().strftime("%Y%m%d")
        m = re.fullmatch(re.escape(prefix) + d + r"-(\d+)", str(ref).strip())
        if m:
            key = f"{kind}:{d}"
            self.data["counters"][key] = max(self.data["counters"].get(key, 0), int(m.group(1)))
        self.add_client(client, contact, site)
        self.log(kind, client, contact, ref, path)
        self.save()

    def open_root(self):
        try:
            os.startfile(self.root)
        except Exception:
            messagebox.showinfo("Projects folder", self.root)


STORE = Store()


def write_pdf(path, story):
    try:
        build_pdf(path, story)
        return True
    except PermissionError:
        messagebox.showerror("Cannot save", f"Close the file if it is open in another program:\n{path}")
    except Exception as e:
        messagebox.showerror("PDF error", str(e))
    return False


class DocMixin:
    """Gives a tab: auto folders, auto ref numbers, local data saving."""
    KIND = ""

    def target(self, ext=".pdf"):
        return STORE.prepare(self.KIND, self.client.get(), self.contact.get(), self.ref.get(), ext)

    def set_ref(self, value):
        self.ref.delete(0, "end")
        self.ref.insert(0, value)

    def finish(self, path):
        STORE.record(self.KIND, self.client.get(), self.contact.get(), self.ref.get(), self.site.get(), path)
        self.set_ref(STORE.next_ref(self.KIND))
        done(path)


C_STRIP = colors.HexColor("#f7f9fb")
C_BAND = colors.HexColor("#1a1d24")
C_LIGHT = colors.HexColor("#eef3f8")
C_LINE = colors.HexColor("#b8c4d0")
C_GREY = colors.HexColor("#6b7380")
C_TEXT = colors.HexColor("#1f2937")
CW = 182 * mm


def esc(x):
    return xml_escape(str(x))


def styles():
    ss = getSampleStyleSheet()
    base = dict(parent=ss["Normal"], fontName=FONT, textColor=C_TEXT)
    return {
        "body": ParagraphStyle("b", fontSize=9.5, leading=15, **base),
        "meta": ParagraphStyle("m", fontSize=8.5, leading=17, **base),
        "metaR": ParagraphStyle("mr", fontSize=8.5, leading=17, alignment=TA_RIGHT, **base),
        "cell": ParagraphStyle("c", fontSize=8.5, leading=11.5, **base),
        "cellB": ParagraphStyle("cb", fontSize=8.5, leading=11.5, **dict(base, fontName=FONT_B)),
        "title": ParagraphStyle("t", fontSize=12.5, leading=17, alignment=TA_CENTER, spaceAfter=2,
                                **dict(base, fontName=FONT_B)),
        "subt": ParagraphStyle("s", fontSize=9, leading=13, alignment=TA_CENTER,
                               **dict(base, fontName=FONT_B, textColor=C_GREY)),
        "right": ParagraphStyle("r", fontSize=9.5, leading=15, alignment=TA_RIGHT, **base),
        "rightB": ParagraphStyle("rb", fontSize=9.5, leading=15, alignment=TA_RIGHT, **dict(base, fontName=FONT_B)),
    }


class NumberedCanvas(canvas_mod.Canvas):
    """Draws the TRIKON header band, contact strip and footer with 'Page X of N'."""

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self._saved = []

    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        n = len(self._saved)
        for st in self._saved:
            self.__dict__.update(st)
            self.chrome(n)
            super().showPage()
        super().save()

    def chrome(self, total):
        w, h = A4
        L, R = 14 * mm, w - 14 * mm
        # header band + gold rule
        self.setFillColor(C_BAND)
        self.rect(0, h - 40 * mm, w, 40 * mm, stroke=0, fill=1)
        self.setFillColor(C_GOLD)
        self.rect(0, h - 41.4 * mm, w, 1.4 * mm, stroke=0, fill=1)
        x = L
        if os.path.exists(LOGO):
            try:
                self.drawImage(LOGO, L, h - 33 * mm, width=22 * mm, height=22 * mm,
                               preserveAspectRatio=True, mask="auto")
                x = L + 25 * mm
            except Exception:
                pass
        self.setFillColor(C_GOLD)
        self.setFont(FONT_B, 21)
        self.drawString(x, h - 15 * mm, FIRM["name"])
        self.setFillColor(colors.white)
        self.setFont(FONT_B, 10.5)
        self.drawString(x, h - 22 * mm, FIRM["sub"])
        self.setFillColor(colors.HexColor("#b4bcc8"))
        self.setFont(FONT_B, 6.5)
        self.drawString(x, h - 27 * mm, FIRM["tag"])
        self.setFillColor(C_GOLD)
        self.setFont(FONT_B, 11)
        self.drawRightString(R, h - 14 * mm, FIRM["owner"])
        self.setFillColor(colors.white)
        self.setFont(FONT_B, 8)
        self.drawRightString(R, h - 20 * mm, FIRM["role"])
        self.setFillColor(colors.HexColor("#b4bcc8"))
        self.setFont(FONT, 6.5)
        self.drawRightString(R, h - 25 * mm, FIRM["quals"])
        # contact strip
        top = h - 41.4 * mm
        self.setFillColor(C_STRIP)
        self.rect(0, top - 10 * mm, w, 10 * mm, stroke=0, fill=1)
        self.setStrokeColor(C_LINE)
        self.setLineWidth(0.5)
        self.line(0, top - 10 * mm, w, top - 10 * mm)
        y = top - 6 * mm
        for lab, val, px in [("Phone:", FIRM["phone"], L + 4 * mm), ("Email:", FIRM["email"], w * 0.38),
                             ("Location:", FIRM["city"], w * 0.70)]:
            self.setFillColor(C_GOLD)
            self.setFont(FONT_B, 7)
            self.drawString(px, y, lab)
            self.setFillColor(C_TEXT)
            self.setFont(FONT, 7)
            self.drawString(px + self.stringWidth(lab, FONT_B, 7) + 4 * mm, y, val)
        # footer
        self.setStrokeColor(C_GOLD)
        self.setLineWidth(1)
        self.line(0, 20 * mm, w, 20 * mm)
        self.setFillColor(C_GOLD)
        self.setFont(FONT_B, 6.5)
        self.drawString(L, 14 * mm, "Office Address:")
        self.setFillColor(C_TEXT)
        self.setFont(FONT, 6.5)
        self.drawString(L + 22 * mm, 14 * mm, "Pl. No. 58, Chhatrapati Sambhaji Maharaj Nagar, Near New Kane Hospital,")
        self.drawString(L + 22 * mm, 10.5 * mm, "Thakare Chowk, Pandharpur, Dist. Solapur - 413304")
        self.setFillColor(C_GOLD)
        self.setFont(FONT_B, 6.5)
        self.drawRightString(R, 14 * mm, "BUILDING THE FUTURE WITH PRECISION")
        self.setFillColor(C_GREY)
        self.setFont(FONT, 6.5)
        self.drawRightString(R, 10.5 * mm, f"Page {self._pageNumber} of {total}")


def build_pdf(path, story):
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=56 * mm, bottomMargin=26 * mm)
    doc.build(story, canvasmaker=NumberedCanvas)


def meta_row(st, left, right):
    t = Table([[Paragraph("<br/>".join(left), st["meta"]), Paragraph("<br/>".join(right), st["metaR"])]],
              colWidths=[112 * mm, 70 * mm])
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                           ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                           ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    return t


def sub_bar(st, text):
    t = Table([[Paragraph(f"<b>SUB: {esc(text)}</b>", st["body"])]], colWidths=[CW])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), C_LIGHT), ("LEFTPADDING", (0, 0), (-1, -1), 8),
                           ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    return t


def signature_block(st):
    return [Spacer(1, 26), Paragraph("<b>FOR TRIKON CONSTRUCTION & DESIGNS</b>", st["right"]), Spacer(1, 34),
            Paragraph(f'<font color="{GOLD}"><b>{FIRM["owner_full"]}</b></font>', st["right"]),
            Paragraph('<font color="#6b7380" size="8">Authorized Signatory / Proprietor</font>', st["right"])]


def two_sign(st, left_head, left_name, right_head, right_name, right_align=False):
    ra = st["right"] if right_align else st["body"]
    left = [Paragraph(f"<b>{left_head}</b>", st["body"]), Spacer(1, 42),
            Paragraph(f"Signature: ____________________<br/>{esc(left_name)}", st["cell"])]
    right = [Paragraph(f"<b>{right_head}</b>", ra), Spacer(1, 42),
             Paragraph(f"Signature: ____________________<br/>{right_name}",
                       ParagraphStyle("x", parent=st["cell"], alignment=TA_RIGHT if right_align else 0))]
    t = Table([[left, right]], colWidths=[91 * mm, 91 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    return [Spacer(1, 26), t]


def grid_style():
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), C_BAND), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), FONT_B), ("FONTNAME", (0, 1), (-1, -1), FONT),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5), ("GRID", (0, 0), (-1, -1), 0.5, C_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("TEXTCOLOR", (0, 1), (-1, -1), C_TEXT),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ])


def done(path):
    if messagebox.askyesno("Saved", f"Saved:\n{path}\n\nOpen it now?"):
        try:
            os.startfile(path)
        except Exception:
            pass


# -------------------------------------------------------------- UI helpers
def edit_dialog(parent, title, fields, values=None):
    """fields: [(label, key, 'entry'|'text')]. Returns dict or None."""
    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent.winfo_toplevel())
    win.grab_set()
    win.configure(bg="white", padx=14, pady=12)
    widgets, result = {}, {}
    for i, (label, key, kind) in enumerate(fields):
        tk.Label(win, text=label, bg="white", font=("Segoe UI", 9, "bold")).grid(row=i, column=0, sticky="nw", pady=4)
        if kind == "text":
            w = tk.Text(win, width=60, height=4, font=("Segoe UI", 9))
            w.insert("1.0", (values or {}).get(key, ""))
        else:
            w = tk.Entry(win, width=62, font=("Segoe UI", 9))
            w.insert(0, (values or {}).get(key, ""))
        w.grid(row=i, column=1, padx=8, pady=4)
        widgets[key] = (w, kind)

    def ok():
        for k, (w, kind) in widgets.items():
            result[k] = (w.get("1.0", "end-1c") if kind == "text" else w.get()).strip()
        win.destroy()

    def cancel():
        result.clear()
        win.destroy()

    bar = tk.Frame(win, bg="white")
    bar.grid(row=len(fields), column=0, columnspan=2, pady=(10, 0), sticky="e")
    tk.Button(bar, text="Save", bg=GOLD, relief="flat", width=10, command=ok).pack(side="left", padx=4)
    tk.Button(bar, text="Cancel", width=10, relief="flat", command=cancel).pack(side="left")
    win.wait_window()
    return dict(result) or None


def labeled_entry(parent, text, width, row, col, default=""):
    tk.Label(parent, text=text, bg="white", font=("Segoe UI", 8, "bold")).grid(row=row, column=col, sticky="e", padx=(12, 4), pady=3)
    e = tk.Entry(parent, width=width, font=("Segoe UI", 8), relief="solid", bd=1)
    e.insert(0, default)
    e.grid(row=row, column=col + 1, sticky="w", pady=3)
    return e


def _fill(entry, text):
    entry.delete(0, "end")
    entry.insert(0, text)


def client_picker(parent, client_e, contact_e, site_e):
    cb = ttk.Combobox(parent, width=44, state="readonly")

    def refresh():
        cb.config(values=STORE.client_keys())

    def pick(_e=None):
        c = STORE.data["clients"].get(cb.get())
        if c:
            _fill(client_e, c.get("client", ""))
            _fill(contact_e, c.get("contact", ""))
            _fill(site_e, c.get("site", ""))

    cb.config(postcommand=refresh)
    cb.bind("<<ComboboxSelected>>", pick)
    return cb


def labeled_picker(parent, row, col, client_e, contact_e, site_e, span=3):
    tk.Label(parent, text="Saved Client:", bg="white", font=("Segoe UI", 8, "bold")).grid(
        row=row, column=col, sticky="e", padx=(12, 4), pady=3)
    client_picker(parent, client_e, contact_e, site_e).grid(row=row, column=col + 1, columnspan=span, sticky="w", pady=3)


class TablePanel(tk.Frame):
    """Treeview + side buttons (Add / Edit / Delete / Reset or Clear)."""

    def __init__(self, parent, columns, widths, fields, add_label, edit_label, defaults=None,
                 on_change=None, anchors=None, clear_mode="reset"):
        super().__init__(parent, bg="white")
        self.columns, self.fields, self.defaults, self.on_change = columns, fields, defaults, on_change
        self.keys = [f[1] for f in fields]
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="browse")
        for c, w in zip(columns, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor=(anchors or {}).get(c, "w"), stretch=True)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        side = tk.Frame(self, bg="white")
        side.pack(side="left", fill="y", padx=6)
        for txt, cmd, bg in [(f"+ {add_label}", self.add, GOLD), (edit_label, self.edit, "#eeeeee"),
                             ("Delete", self.delete, "#eeeeee"),
                             ("Reset Default" if clear_mode == "reset" else "Clear All", self.clear, "#eeeeee")]:
            tk.Button(side, text=txt, bg=bg, relief="flat", width=14, font=("Segoe UI", 8, "bold"),
                      command=cmd).pack(pady=3, ipady=3)
        self.tree.bind("<Double-1>", lambda e: self.edit())
        self.clear_mode = clear_mode
        if defaults:
            self.load(defaults)

    def row_values(self, d):
        return [d[k] for k in self.keys]

    def load(self, rows):
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", values=self.row_values(r))
        self.renumber()

    def rows(self):
        return [dict(zip(self.keys, self.tree.item(i, "values"))) for i in self.tree.get_children()]

    def renumber(self):
        if self.keys and self.keys[0] == "sr":
            for n, i in enumerate(self.tree.get_children(), 1):
                v = list(self.tree.item(i, "values"))
                v[0] = n
                self.tree.item(i, values=v)
        if self.on_change:
            self.on_change()

    def add(self):
        d = edit_dialog(self, "Add", [f for f in self.fields if f[1] != "sr" and f[2] != "auto"])
        if d:
            d = self.finish(d)
            if d:
                self.tree.insert("", "end", values=self.row_values(d))
                self.renumber()

    def edit(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showinfo("Edit", "Select a row first.")
        cur = dict(zip(self.keys, self.tree.item(sel[0], "values")))
        d = edit_dialog(self, "Edit", [f for f in self.fields if f[1] != "sr" and f[2] != "auto"], cur)
        if d:
            cur.update(d)
            cur = self.finish(cur)
            if cur:
                self.tree.item(sel[0], values=self.row_values(cur))
                self.renumber()

    def finish(self, d):
        d.setdefault("sr", "")
        return d

    def delete(self):
        for s in self.tree.selection():
            self.tree.delete(s)
        self.renumber()

    def clear(self):
        if self.clear_mode == "reset" and self.defaults:
            self.load(self.defaults)
        else:
            if messagebox.askyesno("Clear", "Remove all items?"):
                self.tree.delete(*self.tree.get_children())
                self.renumber()


# ------------------------------------------------------------ LETTERHEADS
LETTER_TEMPLATES = {
    "Quotation Submission": ("Submission of Detailed Quotation & Technical Proposal",
        "We are pleased to submit our comprehensive quotation and technical proposal for the proposed construction and finishing works at the designated project site.\n\n"
        "Our estimate is prepared strictly adhering to standard structural and architectural codes, incorporating premium quality raw materials and skilled workmanship to ensure structural durability and fine aesthetics.\n\n"
        "Please find the detailed Bill of Quantities (BOQ) and itemized breakdown attached for your review. We remain available for technical discussions, revisions, or on-site walkthroughs at your convenience."),
    "Work Order Acceptance": ("Acceptance of Work Order",
        "We acknowledge receipt of your work order and are pleased to confirm our acceptance of the scope, terms and conditions mentioned therein.\n\n"
        "We will mobilise manpower and materials at the earliest and keep you informed of progress at every milestone."),
    "Payment Request": ("Request for Release of Running Bill Payment",
        "This is to kindly request release of the payment against the running bill submitted for the work executed as per the agreed milestone.\n\n"
        "We request you to process the same at the earliest as per the payment terms of the agreement."),
}


class LetterTab(DocMixin, tk.Frame):
    KIND = "letter"

    def __init__(self, parent):
        super().__init__(parent, bg="white")
        side = tk.Frame(self, bg=NAVY, width=120)
        side.pack(side="left", fill="y", padx=(10, 0), pady=10)
        for txt, cmd in [("SAVE PDF", self.save_pdf), ("TEMPLATES", self.templates),
                         ("CLEAR", self.clear), ("OPEN FOLDER", STORE.open_root)]:
            tk.Button(side, text=txt, bg=NAVY, fg="white", activebackground=GOLD, relief="flat",
                      font=("Segoe UI", 8, "bold"), width=14, command=cmd).pack(pady=12, padx=8)
        form = tk.Frame(self, bg="white")
        form.pack(side="left", fill="both", expand=True, padx=20, pady=14)
        top = tk.Frame(form, bg="white")
        top.pack(fill="x")
        tk.Label(top, text="Ref. No.:", bg="white", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.ref = tk.Entry(top, width=26, relief="solid", bd=1)
        self.ref.pack(side="left", padx=6)
        self.date = tk.Entry(top, width=14, relief="solid", bd=1)
        self.date.pack(side="right")
        tk.Label(top, text="Date:", bg="white", font=("Segoe UI", 8, "bold")).pack(side="right", padx=4)
        tk.Label(form, text="To,", bg="white", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(10, 2))
        self.to1 = tk.Entry(form, relief="solid", bd=1)
        self.to1.pack(fill="x", pady=1)
        self.to2 = tk.Entry(form, relief="solid", bd=1)
        self.to2.pack(fill="x", pady=1)
        crow = tk.Frame(form, bg="white")
        crow.pack(fill="x", pady=(4, 0))
        tk.Label(crow, text="Contact No.:", bg="white", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.contact = tk.Entry(crow, width=18, relief="solid", bd=1)
        self.contact.pack(side="left", padx=6)
        tk.Label(crow, text="Saved Client:", bg="white", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(16, 4))
        self.client, self.site = self.to1, self.to2
        client_picker(crow, self.to1, self.contact, self.to2).pack(side="left")
        sub = tk.Frame(form, bg="white")
        sub.pack(fill="x", pady=8)
        tk.Label(sub, text="Sub:", bg="white", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.sub = tk.Entry(sub, relief="solid", bd=1, font=("Segoe UI", 9, "bold"))
        self.sub.pack(side="left", fill="x", expand=True, padx=6)
        self.body = tk.Text(form, relief="solid", bd=1, font=("Segoe UI", 9), wrap="word")
        self.body.pack(fill="both", expand=True)
        tk.Label(form, text="FOR TRIKON CONSTRUCTION & DESIGNS", bg="white", font=("Segoe UI", 8, "bold")).pack(anchor="e", pady=(10, 20))
        tk.Label(form, text=FIRM["owner_full"], bg="white", fg=GOLD, font=("Segoe UI", 9, "bold")).pack(anchor="e")
        tk.Label(form, text="Authorized Signatory / Proprietor", bg="white", font=("Segoe UI", 8)).pack(anchor="e")
        self.clear()
        self.apply_template("Quotation Submission")

    def clear(self):
        for e, v in [(self.ref, STORE.next_ref("letter")), (self.contact, ""), (self.date, today()), (self.to1, "M/s "),
                     (self.to2, "Site Location / Address Line"), (self.sub, "")]:
            e.delete(0, "end")
            e.insert(0, v)
        self.body.delete("1.0", "end")

    def apply_template(self, name):
        s, b = LETTER_TEMPLATES[name]
        self.sub.delete(0, "end")
        self.sub.insert(0, s)
        self.body.delete("1.0", "end")
        self.body.insert("1.0", b)

    def templates(self):
        m = tk.Menu(self, tearoff=0)
        for n in LETTER_TEMPLATES:
            m.add_command(label=n, command=lambda n=n: self.apply_template(n))
        m.tk_popup(self.winfo_pointerx(), self.winfo_pointery())

    def story(self):
        st = styles()
        s = [meta_row(st, [f"Ref. No.: {esc(self.ref.get())}"], [f"Date: {esc(self.date.get())}"]),
             Spacer(1, 12), Paragraph("<b>To,</b>", st["body"]),
             Paragraph(f"<b>{esc(self.to1.get())}</b>", st["body"]), Paragraph(esc(self.to2.get()), st["body"]),
             Spacer(1, 10), sub_bar(st, self.sub.get().upper()), Spacer(1, 12)]
        for para in self.body.get("1.0", "end-1c").split("\n\n"):
            if para.strip():
                s += [Paragraph(esc(para.strip()).replace("\n", "<br/>"), st["body"]), Spacer(1, 8)]
        return s + signature_block(st)

    def save_pdf(self):
        p = self.target()
        if p and write_pdf(p, self.story()):
            self.finish(p)


# -------------------------------------------------------------- BOQ
class BOQTab(DocMixin, tk.Frame):
    KIND = "boq"

    def __init__(self, parent):
        super().__init__(parent, bg="white")
        self.dirty, self.last_ref = True, ""
        hdr = tk.Frame(self, bg="white", bd=1, relief="solid")
        hdr.pack(fill="x", padx=10, pady=8)
        self.client = labeled_entry(hdr, "Client Name:", 24, 0, 0, "M/s Private Residence")
        self.site = labeled_entry(hdr, "Project Site:", 24, 0, 2, "Pandharpur Project Site")
        self.ref = labeled_entry(hdr, "Estimate Ref:", 20, 0, 4, STORE.next_ref("boq"))
        self.contact = labeled_entry(hdr, "Contact No.:", 24, 1, 0, "")
        labeled_picker(hdr, 1, 2, self.client, self.contact, self.site)
        defaults = [
            dict(sr=1, desc="Earthwork excavation in all types of soil including disposal", qty="450", unit="cu.m", rate="185", amt=""),
            dict(sr=2, desc="P.C.C. (1:4:8) for foundation bed including curing & ramming", qty="35", unit="cu.m", rate="4800", amt=""),
            dict(sr=3, desc="R.C.C. M25 grade for Columns, Beams & Slabs with formwork", qty="62", unit="cu.m", rate="8200", amt=""),
            dict(sr=4, desc="High Yield TMT Steel Reinforcement Fe550D cutting & binding", qty="5400", unit="kg", rate="72", amt=""),
            dict(sr=5, desc="Red Brickwork masonry 6 inch thick in cement mortar (1:6)", qty="1850", unit="sq.ft", rate="145", amt=""),
        ]
        fields = [("Sr.", "sr", "auto"), ("Item Description / Specification", "desc", "text"),
                  ("Quantity", "qty", "entry"), ("Unit", "unit", "entry"), ("Rate (\u20b9)", "rate", "entry"),
                  ("Amount (\u20b9)", "amt", "auto")]
        self.panel = TablePanel(self, ["Sr.", "Item Description / Specification", "Quantity", "Unit", "Rate (\u20b9)", "Amount (\u20b9)"],
                                [40, 420, 80, 70, 100, 110], fields, "Add Item", "Edit Item",
                                on_change=self.totals, clear_mode="clear",
                                anchors={"Sr.": "center", "Quantity": "center", "Unit": "center", "Rate (\u20b9)": "e", "Amount (\u20b9)": "e"})
        self.panel.finish = self.finish_row
        self.panel.pack(fill="both", expand=True, padx=10)
        self.panel.load([self.finish_row(d) for d in defaults])
        bar = tk.Frame(self, bg="white", bd=1, relief="solid")
        bar.pack(fill="x", padx=10, pady=8)
        tk.Label(bar, text="Overheads/Profit (%):", bg="white", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(12, 4), pady=10)
        self.oh = tk.Entry(bar, width=6, relief="solid", bd=1)
        self.oh.insert(0, "10")
        self.oh.pack(side="left")
        tk.Label(bar, text="GST (%):", bg="white", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(16, 4))
        self.gst = tk.Entry(bar, width=6, relief="solid", bd=1)
        self.gst.insert(0, "18")
        self.gst.pack(side="left")
        for e in (self.oh, self.gst):
            e.bind("<KeyRelease>", lambda ev: self.totals())
        self.grand_lbl = tk.Label(bar, bg="white", font=("Segoe UI", 11, "bold"))
        self.grand_lbl.pack(side="right", padx=12)
        self.sub_lbl = tk.Label(bar, bg="white", font=("Segoe UI", 8, "bold"))
        self.sub_lbl.pack(side="right", padx=12)
        tk.Button(bar, text="Export to Excel", relief="solid", bd=1, command=self.excel).pack(side="right", padx=4)
        tk.Button(bar, text="Generate BOQ PDF", bg=GOLD, relief="flat", font=("Segoe UI", 8, "bold"),
                  command=self.pdf).pack(side="right", padx=4)
        self.totals()

    def finish_row(self, d):
        try:
            q, r = float(str(d["qty"]).replace(",", "")), float(str(d["rate"]).replace(",", ""))
            d["amt"], d["rate"] = money(q * r), money(r)
        except ValueError:
            messagebox.showerror("Invalid", "Quantity and Rate must be numbers.")
            return None
        d.setdefault("sr", "")
        return d

    def num(self, s):
        return float(str(s).replace(",", "") or 0)

    def calc(self):
        sub = sum(self.num(r["amt"]) for r in self.panel.rows())
        try:
            oh, gst = self.num(self.oh.get()), self.num(self.gst.get())
        except ValueError:
            oh = gst = 0
        ohv = sub * oh / 100
        gstv = (sub + ohv) * gst / 100
        return sub, oh, ohv, gst, gstv, sub + ohv + gstv

    def finish(self, p):
        self.last_ref, self.dirty = self.ref.get(), False
        super().finish(p)

    def totals(self):
        self.dirty = True
        if not hasattr(self, "grand_lbl"):
            return
        sub, _, _, _, _, grand = self.calc()
        self.sub_lbl.config(text=f"Subtotal: {RS}{money(sub)}")
        self.grand_lbl.config(text=f"Grand Total: {RS}{money(grand)}")

    def pdf(self):
        p = self.target()
        if not p:
            return
        st = styles()
        sub, oh, ohv, gst, gstv, grand = self.calc()
        data = [["Sr.", "Item Description", "Qty", "Unit", f"Rate ({RS})", f"Amount ({RS})"]]
        for r in self.panel.rows():
            data.append([r["sr"], Paragraph(esc(r["desc"]), st["cell"]), r["qty"], r["unit"], r["rate"], r["amt"]])
        n = len(data)
        data += [["", "SUBTOTAL", "", "", "", f"{RS}{money(sub)}"],
                 ["", f"Overhead & Contractor Margin ({oh:.1f}%)", "", "", "", f"{RS}{money(ohv)}"],
                 ["", f"GST ({gst:.1f}%)", "", "", "", f"{RS}{money(gstv)}"],
                 ["", "GRAND TOTAL", "", "", "", f"{RS}{money(grand)}"]]
        t = Table(data, colWidths=[12 * mm, 86 * mm, 18 * mm, 16 * mm, 24 * mm, 26 * mm], repeatRows=1)
        ts = grid_style()
        ts.add("ALIGN", (0, 1), (0, n - 1), "CENTER")
        ts.add("ALIGN", (2, 1), (3, n - 1), "CENTER")
        ts.add("ALIGN", (4, 1), (-1, -1), "RIGHT")
        for r in range(n, n + 4):
            ts.add("SPAN", (1, r), (4, r))
        ts.add("ALIGN", (1, n), (1, n), "RIGHT")
        ts.add("FONTNAME", (1, n), (-1, n), FONT_B)
        ts.add("BACKGROUND", (0, n + 3), (-1, n + 3), C_LIGHT)
        ts.add("FONTNAME", (1, n + 3), (-1, n + 3), FONT_B)
        ts.add("TEXTCOLOR", (1, n + 3), (-1, n + 3), C_GOLD)
        ts.add("FONTNAME", (5, n), (5, n), FONT_B)
        t.setStyle(ts)
        story = [meta_row(st, [f"Client: {esc(self.client.get())}", f"Site: {esc(self.site.get())}"],
                          [f"Date: {today()}", f"Estimate Ref: {esc(self.ref.get())}"]), Spacer(1, 8), t]
        if write_pdf(p, story):
            self.finish(p)

    def excel(self):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill
        except ImportError:
            return messagebox.showerror("Missing", "Run: pip install openpyxl")
        ref = self.last_ref if (self.last_ref and not self.dirty) else self.ref.get()
        p = STORE.prepare("boq", self.client.get(), self.contact.get(), ref, ".xlsx")
        if not p:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "BOQ"
        ws.append(["TRIKON CONSTRUCTION & DESIGNS"])
        ws.append([f"Client: {self.client.get()}", f"Site: {self.site.get()}", f"Ref: {ref}"])
        ws.append([])
        ws.append(["Sr.", "Description", "Qty", "Unit", "Rate", "Amount"])
        for c in ws[4]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="161B26")
        for r in self.panel.rows():
            ws.append([int(r["sr"]), r["desc"], self.num(r["qty"]), r["unit"], self.num(r["rate"]), self.num(r["amt"])])
        sub, oh, ohv, gst, gstv, grand = self.calc()
        for lab, v in [("Subtotal", sub), (f"Overheads/Profit {oh:g}%", ohv), (f"GST {gst:g}%", gstv), ("GRAND TOTAL", grand)]:
            ws.append(["", lab, "", "", "", v])
        ws["A1"].font = Font(bold=True, size=14)
        ws.column_dimensions["B"].width = 70
        for col in "CDEF":
            ws.column_dimensions[col].width = 14
        try:
            wb.save(p)
        except Exception as e:
            return messagebox.showerror("Excel error", str(e))
        STORE.log("boq", self.client.get(), self.contact.get(), ref, p)
        done(p)


# ---------------------------------------------------------- CONTRACT
DEFAULT_CLAUSES = (
    "1. SCOPE OF EXECUTION: The Contractor agrees to execute and complete all civil, structural, and finishing works as per architectural drawings and approved BOQ specifications.\n\n"
    "2. TIMELINE & DELAYS: The project shall be completed within the stipulated duration from site handover. Delays arising from force majeure or client-requested changes shall extend the timeline proportionately.\n\n"
    "3. STAGE PAYMENT SCHEDULE: Payments shall be released strictly in accordance with milestone completion (Plinth, Slabs, Masonry, Finishing, and Handover). Running bills must be cleared within 7 days of submission.\n\n"
    "4. VARIATIONS & EXTRA ITEMS: Any work executed beyond the agreed BOQ scope shall be treated as an extra item and billed separately based on actual market rates + contractor margin.\n\n"
    "5. DEFECT LIABILITY PERIOD: The Contractor provides a 12-month Defect Liability Period from formal handover to rectify any structural or execution defects free of charge (excluding normal wear & tear)."
)


class ContractTab(DocMixin, tk.Frame):
    KIND = "contract"

    def __init__(self, parent):
        super().__init__(parent, bg="white")
        hdr = tk.Frame(self, bg="white", bd=1, relief="solid")
        hdr.pack(fill="x", padx=10, pady=8)
        self.client = labeled_entry(hdr, "Client (First Party):", 22, 0, 0, "Mr. / Mrs. Client Name")
        self.site = labeled_entry(hdr, "Project Site:", 26, 0, 2, "Plot No., Location, Pandharpur")
        self.ref = labeled_entry(hdr, "Contract Ref:", 22, 0, 4, STORE.next_ref("contract"))
        self.area = labeled_entry(hdr, "Built-up Area (Sq.Ft):", 22, 1, 0, "1500")
        self.rate = labeled_entry(hdr, "Rate / Sq.Ft:", 26, 1, 2, "1800")
        self.sum = labeled_entry(hdr, "Agreed Sum (\u20b9):", 22, 1, 4, "2,700,000.00")
        self.dur = labeled_entry(hdr, "Duration:", 22, 2, 0, "8 Calendar Months")
        self.dt = labeled_entry(hdr, "Agreement Date:", 26, 2, 2, today())
        self.contact = labeled_entry(hdr, "Contact No.:", 22, 2, 4, "")
        labeled_picker(hdr, 3, 0, self.client, self.contact, self.site)
        for e in (self.area, self.rate):
            e.bind("<KeyRelease>", lambda ev: self.recalc())
        tk.Label(self, text="Contractual Clauses & Legal Terms:", bg="white", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=12)
        self.text = tk.Text(self, font=("Segoe UI", 9), wrap="word", relief="solid", bd=1)
        self.text.pack(fill="both", expand=True, padx=10, pady=4)
        self.text.insert("1.0", DEFAULT_CLAUSES)
        bar = tk.Frame(self, bg="white", bd=1, relief="solid")
        bar.pack(fill="x", padx=10, pady=8)
        tk.Button(bar, text="Reset Clauses to Default", relief="solid", bd=1, command=self.reset).pack(side="left", padx=8, pady=8)
        tk.Button(bar, text="Generate Contract Agreement PDF", bg=GOLD, relief="flat", font=("Segoe UI", 8, "bold"),
                  command=self.pdf).pack(side="right", padx=8)

    def recalc(self):
        try:
            v = float(self.area.get().replace(",", "")) * float(self.rate.get().replace(",", ""))
            self.sum.delete(0, "end")
            self.sum.insert(0, money(v))
        except ValueError:
            pass

    def reset(self):
        self.text.delete("1.0", "end")
        self.text.insert("1.0", DEFAULT_CLAUSES)

    def pdf(self):
        p = self.target()
        if not p:
            return
        st = styles()
        cb = ParagraphStyle("cbx", parent=st["cell"], fontName=FONT_B, fontSize=8.5)
        info = [[Paragraph(f"Project Location: {esc(self.site.get())}", cb), Paragraph(f"Built-up Area: {esc(self.area.get())} sq.ft", cb)],
                [Paragraph(f"Contract Rate: {RS} {esc(self.rate.get())} / sq.ft", cb), Paragraph(f"Agreed Contract Sum: {RS} {esc(self.sum.get())}/-", cb)],
                [Paragraph(f"Project Duration: {esc(self.dur.get())}", cb), Paragraph("Defect Liability: 12 Months", cb)]]
        t = Table(info, colWidths=[91 * mm, 91 * mm])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), C_STRIP), ("GRID", (0, 0), (-1, -1), 0.5, C_LINE),
                               ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
        story = [Paragraph("CONSTRUCTION EXECUTION AGREEMENT", st["title"]),
                 Paragraph("ARTICLES OF AGREEMENT & CONTRACT SPECIFICATIONS", st["subt"]), Spacer(1, 12),
                 Paragraph(f"This Agreement is executed on {esc(self.dt.get())} (Ref: {esc(self.ref.get())}) by and between "
                           f"{esc(self.client.get())} (hereinafter designated as the 'FIRST PARTY / CLIENT') and "
                           f"TRIKON CONSTRUCTION &amp; DESIGNS, represented by {FIRM['owner']} "
                           f"(hereinafter designated as the 'SECOND PARTY / CONTRACTOR').", st["body"]),
                 Spacer(1, 8), t, Spacer(1, 10)]
        for para in self.text.get("1.0", "end-1c").split("\n\n"):
            if para.strip():
                story += [Paragraph(esc(para.strip()), st["body"]), Spacer(1, 6)]
        story += two_sign(st, "FIRST PARTY (CLIENT)", self.client.get(), "SECOND PARTY (CONTRACTOR)",
                          f"{FIRM['owner_full']}<br/>Authorized Signatory / Proprietor")
        if write_pdf(p, story):
            self.finish(p)



# ----------------------------------------------------- MATERIAL & SPECS
DEFAULT_SPECS = [
    ("Civil Structure", "Cement", "UltraTech / ACC / Ambuja (53 / 43 Grade OPC & PPC)"),
    ("Civil Structure", "Reinforcement Steel", "Tata Tiscon / Jindal Panther / JSW Fe550D TMT Bars"),
    ("Civil Structure", "Sand / Aggregates", "Washed Crushed M-Sand & Plaster Sand; 20mm & 10mm Basalt Aggregates"),
    ("Civil Structure", "Masonry / Bricks", "Standard Kiln Red Bricks (6\" & 4\") / High-Grade AAC Blocks"),
    ("Finishing", "Flooring & Tiling", "Kajaria / Nitco Vitrified Tiles (1200x600mm Living, 600x600mm Bed/Kitchen)"),
    ("Finishing", "Wall & Ceiling Paint", "Asian Paints Royale (Internal) & Apex Ultima Weather-proof (External)"),
    ("Finishing", "Waterproofing", "Dr. Fixit Pidifin 2K / Fosroc Brushbond for Toilets, Terrace & Sump"),
    ("MEP - Plumbing", "Pipes & Fittings", "Astral / Finolex CPVC (Hot/Cold water) & Supreme SWR Drainage Pipes"),
    ("MEP - Plumbing", "Sanitary & CP Fixtures", "Jaquar / Kohler Diverters, Wall-hung Closets & Chrome-finish Taps"),
    ("MEP - Electrical", "Conduits & Wires", "Polycab / Havells Flame Retardant (FR) Multi-strand Copper Wiring"),
    ("MEP - Electrical", "Modular Switches", "Legrand / Anchor Roma Modular Plates, MCBs & Distribution Boards"),
    ("Joinery & Openings", "Doors & Frames", "Granite/Teak Frames with Marine Grade 32mm Laminated Flush Doors"),
    ("Joinery & Openings", "Windows", "Jindal Anodized 3-Track Aluminum / Domal Series Sliding Windows with Mosquito Mesh"),
]


class SpecsTab(DocMixin, tk.Frame):
    KIND = "specs"

    def __init__(self, parent):
        super().__init__(parent, bg="white")
        hdr = tk.Frame(self, bg="white", bd=1, relief="solid")
        hdr.pack(fill="x", padx=10, pady=8)
        self.client = labeled_entry(hdr, "Project / Client:", 24, 0, 0, "M/s Residential Project")
        self.site = labeled_entry(hdr, "Site Location:", 24, 0, 2, "Pandharpur, Maharashtra")
        self.ref = labeled_entry(hdr, "Document Ref:", 22, 0, 4, STORE.next_ref("specs"))
        self.contact = labeled_entry(hdr, "Contact No.:", 24, 1, 0, "")
        labeled_picker(hdr, 1, 2, self.client, self.contact, self.site)
        self.defaults = [dict(cat=a, mat=b, spec=c) for a, b, c in DEFAULT_SPECS]
        self.panel = TablePanel(self, ["Trade Category", "Construction Material", "Approved Brand / Grade / Technical Specification"],
                                [200, 220, 560],
                                [("Trade Category", "cat", "entry"), ("Construction Material", "mat", "entry"),
                                 ("Approved Brand / Grade / Technical Specification", "spec", "text")],
                                "Add Spec", "Edit Spec", defaults=self.defaults)
        self.panel.pack(fill="both", expand=True, padx=10)
        bar = tk.Frame(self, bg="white", bd=1, relief="solid")
        bar.pack(fill="x", padx=10, pady=8)
        tk.Button(bar, text="Generate Material Specifications PDF", bg=GOLD, relief="flat", font=("Segoe UI", 8, "bold"),
                  command=self.pdf).pack(side="right", padx=8, pady=10)

    def pdf(self):
        p = self.target()
        if not p:
            return
        st = styles()
        data = [["Trade Category", "Material Item", "Approved Brand / Technical Specification"]]
        for r in self.panel.rows():
            data.append([Paragraph(esc(r["cat"]), st["cell"]), Paragraph(esc(r["mat"]), st["cell"]),
                         Paragraph(esc(r["spec"]), st["cell"])])
        t = Table(data, colWidths=[34 * mm, 42 * mm, 106 * mm], repeatRows=1)
        t.setStyle(grid_style())
        story = [meta_row(st, [f"Project / Client: {esc(self.client.get())}", f"Site Location: {esc(self.site.get())}"],
                          [f"Date: {today()}", f"Spec Ref: {esc(self.ref.get())}"]), Spacer(1, 8), t]
        if write_pdf(p, story):
            self.finish(p)



# --------------------------------------------------------- PROJECT SCOPE
DEFAULT_SCOPE = [
    ("1. Architectural & Engineering", "Preparation of architectural 2D floor plans, 3D exterior elevations, structural engineering drawings, and site layout demarcation."),
    ("2. Substructure & Foundation", "Site clearance, excavation in soil/soft rock, PCC foundation bed (1:4:8), RCC footings, columns up to plinth level, and plinth beam casting."),
    ("3. Superstructure RCC", "RCC columns, beams, staircase, and slab casting using ready-mix or machine-mixed M25 concrete with Fe550D TMT steel reinforcement."),
    ("4. Masonry & Plastering", "External 6\" and internal 4\" red brick masonry in 1:6 cement mortar. Double-coat sand-faced external plaster and smooth internal gypsum/plaster."),
    ("5. Waterproofing Protection", "Chemical polymer waterproofing in toilet sunken slabs, water tank, and terrace slab with brickbat coba water-tight membrane."),
    ("6. Plumbing & Drainage", "Concealed CPVC piping for hot/cold water supply, SWR drainage piping lines, inspection chambers, and provision for overhead water tank."),
    ("7. Electrical Conduiting", "Concealed fire-resistant electrical PVC conduits, wiring, distribution board with MCB protection, and modular switch box provision."),
    ("8. Exclusions / Client Scope", "Electricity & water supply connection for construction work, municipal approval fees, loose furniture, home automation, and MSEB meter costs."),
]


class ScopeTab(DocMixin, tk.Frame):
    KIND = "scope"

    def __init__(self, parent):
        super().__init__(parent, bg="white")
        hdr = tk.Frame(self, bg="white", bd=1, relief="solid")
        hdr.pack(fill="x", padx=10, pady=8)
        self.client = labeled_entry(hdr, "Project / Client:", 24, 0, 0, "M/s Residential Bungalow")
        self.site = labeled_entry(hdr, "Project Site:", 24, 0, 2, "Pandharpur, Dist. Solapur")
        self.area = labeled_entry(hdr, "Built-up Area:", 14, 0, 4, "2,400 sq.ft")
        self.ref = labeled_entry(hdr, "Scope Ref:", 24, 1, 0, STORE.next_ref("scope"))
        self.contact = labeled_entry(hdr, "Contact No.:", 18, 1, 4, "")
        labeled_picker(hdr, 2, 0, self.client, self.contact, self.site)
        self.dt = labeled_entry(hdr, "Date:", 24, 1, 2, today())
        self.defaults = [dict(phase=a, desc=b) for a, b in DEFAULT_SCOPE]
        self.panel = TablePanel(self, ["Project Phase / Trade", "Inclusions & Scope Boundary Description"], [280, 700],
                                [("Project Phase / Trade", "phase", "entry"), ("Inclusions & Scope Boundary Description", "desc", "text")],
                                "Add Phase", "Edit Phase", defaults=self.defaults)
        self.panel.pack(fill="both", expand=True, padx=10)
        box = tk.Frame(self, bg="white", bd=1, relief="solid")
        box.pack(fill="x", padx=10, pady=8)
        tk.Label(box, text="General Boundary Conditions & Notes:", bg="white", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=6, pady=(6, 0))
        self.notes = tk.Text(box, height=3, font=("Segoe UI", 9), relief="solid", bd=1)
        self.notes.pack(fill="x", padx=6, pady=4)
        self.notes.insert("1.0", "Any changes beyond this agreed scope of work will be billed as variation orders upon mutual written consent.")
        tk.Button(box, text="Generate Project Scope Document PDF", bg=GOLD, relief="flat", font=("Segoe UI", 8, "bold"),
                  command=self.pdf).pack(anchor="e", padx=6, pady=6)

    def pdf(self):
        p = self.target()
        if not p:
            return
        st = styles()
        data = [["Trade Phase", "Inclusions & Boundary Specifications"]]
        for r in self.panel.rows():
            data.append([Paragraph(esc(r["phase"]), st["cellB"]), Paragraph(esc(r["desc"]), st["cell"])])
        t = Table(data, colWidths=[55 * mm, 127 * mm], repeatRows=1)
        t.setStyle(grid_style())
        story = [Paragraph("PROJECT SCOPE OF WORK & EXECUTION BOUNDARIES".replace("&", "&amp;"), st["title"]), Spacer(1, 8),
                 meta_row(st, [f"Client: {esc(self.client.get())}", f"Site Location: {esc(self.site.get())}",
                               f"Proposed Built-up Area: {esc(self.area.get())}"],
                          [f"Date: {esc(self.dt.get())}", f"Scope Ref: {esc(self.ref.get())}"]),
                 Spacer(1, 8), t, Spacer(1, 8),
                 Paragraph(f"<font color=\"#6b7380\">Notes & Conditions: {esc(self.notes.get('1.0', 'end-1c'))}</font>".replace(" & ", " &amp; "), st["cell"])]
        story += two_sign(st, "CLIENT ACCEPTANCE", self.client.get(), "FOR TRIKON CONSTRUCTION &amp; DESIGNS",
                          f"{FIRM['owner_full']}<br/>Authorized Signatory / Proprietor", right_align=True)
        if write_pdf(p, story):
            self.finish(p)



# ------------------------------------------------------------------ MAIN
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TRIKON - Construction Engineering Studio")
        self.geometry("1280x720")
        self.minsize(1000, 600)
        try:
            self.state("zoomed")
        except tk.TclError:
            pass
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook", background=NAVY, borderwidth=0, tabmargins=0)
        style.configure("TNotebook.Tab", background="#232a3a", foreground="white", padding=(16, 8),
                        font=("Segoe UI", 8, "bold"), borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", GOLD)], foreground=[("selected", NAVY)])
        style.configure("Treeview", rowheight=22, font=("Segoe UI", 8))
        style.configure("Treeview.Heading", background=NAVY, foreground="white", font=("Segoe UI", 8, "bold"))
        style.map("Treeview.Heading", background=[("active", NAVY)])

        head = tk.Frame(self, bg=NAVY)
        head.pack(fill="x")
        left = tk.Frame(head, bg=NAVY)
        left.pack(side="left", padx=18, pady=10)
        tk.Label(left, text=FIRM["name"], bg=NAVY, fg=GOLD, font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tk.Label(left, text=FIRM["sub"], bg=NAVY, fg="white", font=("Segoe UI", 8, "bold")).pack(anchor="w")
        tk.Label(left, text=FIRM["tag"], bg=NAVY, fg="#bbbbbb", font=("Segoe UI", 7)).pack(anchor="w")
        right = tk.Frame(head, bg=NAVY)
        right.pack(side="right", padx=18)
        tk.Label(right, text=FIRM["owner"], bg=NAVY, fg=GOLD, font=("Segoe UI", 10, "bold")).pack(anchor="e")
        tk.Label(right, text=FIRM["role"], bg=NAVY, fg="white", font=("Segoe UI", 7, "bold")).pack(anchor="e")
        tk.Label(right, text=FIRM["quals"], bg=NAVY, fg="#bbbbbb", font=("Segoe UI", 7)).pack(anchor="e")
        info = tk.Frame(self, bg="#eef0f4")
        info.pack(fill="x")
        tk.Label(info, text=f"\u260e {FIRM['phone']}   |   \u2709 {FIRM['email']}   |   \u2316 {FIRM['city']}",
                 bg="#eef0f4", font=("Segoe UI", 7)).pack(side="left", padx=18, pady=3)
        tk.Button(info, text="\U0001F4C1 Open Projects Folder", relief="flat", bg="#eef0f4", font=("Segoe UI", 7, "bold"),
                  command=STORE.open_root).pack(side="right", padx=18)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        for name, cls in [("LETTERHEADS", LetterTab), ("BOQ GENERATOR", BOQTab), ("CONTRACT GENERATOR", ContractTab),
                          ("MATERIAL & SPECS", SpecsTab), ("PROJECT SCOPE", ScopeTab)]:
            nb.add(cls(nb), text=name)


if __name__ == "__main__":
    App().mainloop()
