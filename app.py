"""
Invoice Converter - a small desktop app (Mac and Windows) around pdf_to_invoice.py.

Drop a scanned timesheet PDF on the window, confirm the invoice number,
pick where to save, and it writes one Excel invoice per timesheet.
"""
import datetime as dt
import json
import os
import subprocess
import sys
import threading
import traceback

IS_WIN = os.name == "nt"
IS_MAC = sys.platform == "darwin"


def _report_file():
    """Optional file to write test output to (a Windows app has no console)."""
    for flag in ("--selftest", "--convert"):
        if flag in sys.argv:
            rest = sys.argv[sys.argv.index(flag) + 1:]
            for a in rest:
                if a.startswith("--report="):
                    return a.split("=", 1)[1]
    return None


def _out(text):
    path = _report_file()
    if path:
        with open(path, "a", encoding="utf-8") as f:
            f.write(text + "\n")
    if sys.stdout:
        try:
            print(text)
        except Exception:
            pass


# Self-test used by the build: load every library the app needs and exit.
# Catches a broken build (missing or clashing libraries) before it's installed.
if "--selftest" in sys.argv:
    try:
        import ssl, numpy, cv2, pymupdf, pytesseract  # noqa: F401
        import webview  # noqa: F401
        from webview.dom import DOMEventHandler  # noqa: F401
        if IS_MAC:
            import webview.platforms.cocoa  # noqa: F401
        if IS_WIN:
            import webview.platforms.winforms  # noqa: F401
        import pdf_to_invoice  # noqa: F401
        import labels  # noqa: F401
        res = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        for f in ("ui.html", "template.xlsx", "labels_template.docx"):
            if not os.path.exists(os.path.join(res, f)):
                raise FileNotFoundError(f"{f} missing from the app")
        tess = pdf_to_invoice.find_tesseract()
        if not tess:
            raise FileNotFoundError("Tesseract not found")
        pytesseract.pytesseract.tesseract_cmd = tess
        tv = pytesseract.get_tesseract_version()
        _out(f"SELFTEST OK (numpy {numpy.__version__}, opencv {cv2.__version__}, "
             f"pymupdf {pymupdf.__version__}, {ssl.OPENSSL_VERSION}, tesseract {tv} at {tess})")
        sys.exit(0)
    except SystemExit:
        raise
    except BaseException:
        _out(traceback.format_exc())
        _out("SELFTEST FAILED")
        sys.exit(1)

# Headless conversion used by the automated build test:
#   app --convert PDF OUT_DIR INVOICE_NO --report=results.txt
if "--convert" in sys.argv:
    try:
        import pytesseract
        import pdf_to_invoice as _conv
        i = sys.argv.index("--convert")
        pdf, out_dir, inv = sys.argv[i + 1:i + 4]
        pytesseract.pytesseract.tesseract_cmd = _conv.find_tesseract()
        res = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        import labels as _lab
        groups = _conv.read_pdf(pdf, os.path.join(res, "template.xlsx"))
        results = _conv.write_invoices(groups, os.path.join(res, "template.xlsx"), inv, out_dir)
        people = _lab.label_people(groups)
        matches, _ = _lab.match_checks(people, _lab.DEFAULT_CHECKS)
        _lab.build_labels(os.path.join(res, "labels_template.docx"), os.path.join(out_dir, "labels.docx"),
                          people, {m["name"]: m["amount"] for m in matches})
        _out(f"LABELS people={len(people)} checks={len(matches)}")
        for r in results:
            if "error" in r:
                _out(f"RESULT page={r['page']} ERROR {r['error']}")
            else:
                _out(f"RESULT page={r['page']} label={r['label']!r} employees={r['employees']} "
                     f"total={r['total']:.2f} pdf_total={r['pdf_total']} total_ok={r['total_ok']} "
                     f"issues={len(r['issues'])}")
        sys.exit(0)
    except SystemExit:
        raise
    except BaseException:
        _out(traceback.format_exc())
        sys.exit(1)

APP_NAME = "Invoice Converter"
if IS_WIN:
    SUPPORT_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), APP_NAME)
    LOG_FILE = os.path.join(SUPPORT_DIR, "Invoice Converter.log")
else:
    SUPPORT_DIR = os.path.expanduser(f"~/Library/Application Support/{APP_NAME}")
    LOG_FILE = os.path.expanduser("~/Library/Logs/Invoice Converter.log")

# When running as the built app there is no Terminal to show errors, so write
# them to a log file (Mac: ~/Library/Logs, Windows: %APPDATA%\\Invoice Converter)
if getattr(sys, "frozen", False):
    try:
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        _log = open(LOG_FILE, "a", buffering=1, encoding="utf-8")
        sys.stdout = sys.stderr = _log
        print(f"\n=== Invoice Converter started {dt.datetime.now():%Y-%m-%d %H:%M} ===")
    except Exception:
        pass

import webview
from webview.dom import DOMEventHandler

import pytesseract
import pdf_to_invoice as conv
import labels as lab

RES_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
SETTINGS = os.path.join(SUPPORT_DIR, "settings.json")
CHECKS = os.path.join(SUPPORT_DIR, "checks.json")
NAMES = os.path.join(SUPPORT_DIR, "names.json")


def _desktop():
    cands = [os.path.expanduser("~/Desktop")]
    if IS_WIN and os.environ.get("OneDrive"):
        cands.insert(0, os.path.join(os.environ["OneDrive"], "Desktop"))  # OneDrive-synced Desktop
    return next((c for c in cands if os.path.isdir(c)), os.path.expanduser("~"))


DESKTOP = _desktop()


def load_settings():
    try:
        with open(SETTINGS) as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    try:
        os.makedirs(SUPPORT_DIR, exist_ok=True)
        with open(SETTINGS, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def template_path():
    """A template.xlsx placed in Application Support overrides the built-in one
    (so the layout or Employee List can be updated without rebuilding)."""
    custom = os.path.join(SUPPORT_DIR, "template.xlsx")
    return custom if os.path.exists(custom) else os.path.join(RES_DIR, "template.xlsx")


def labels_template_path():
    custom = os.path.join(SUPPORT_DIR, "labels_template.docx")
    return custom if os.path.exists(custom) else os.path.join(RES_DIR, "labels_template.docx")


def load_checks():
    """The check list (who gets part of their pay as a check, and how much).
    Starts from the built-in list and remembers every confirmed change."""
    try:
        with open(CHECKS, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return [dict(e) for e in lab.DEFAULT_CHECKS]


def save_checks(data):
    try:
        os.makedirs(SUPPORT_DIR, exist_ok=True)
        with open(CHECKS, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        traceback.print_exc()


def load_names():
    """Names confirmed in earlier runs ("known") and spelling fixes the user
    made ("fixes": name as the scan read it -> correct name)."""
    try:
        with open(NAMES, encoding="utf-8") as f:
            d = json.load(f)
        return {"known": list(d.get("known", [])), "fixes": dict(d.get("fixes", {}))}
    except Exception:
        return {"known": [], "fixes": {}}


def save_names(d):
    try:
        os.makedirs(SUPPORT_DIR, exist_ok=True)
        with open(NAMES, "w", encoding="utf-8") as f:
            json.dump(d, f, indent=2, ensure_ascii=False)
    except Exception:
        traceback.print_exc()


class Api:
    def __init__(self):
        self._window = None
        self._busy = False
        self._pending = None  # PDF read, waiting for check amounts to be confirmed

    # ---- state -----------------------------------------------------------
    def get_state(self):
        s = load_settings()
        last = s.get("last_invoice_no")
        return {
            "next_invoice_no": conv.next_invoice_no(last, 1) if last else "",
            "tesseract_ok": conv.find_tesseract() is not None,
            "platform": "windows" if IS_WIN else ("mac" if IS_MAC else "other"),
        }

    # ---- file pickers ----------------------------------------------------
    def choose_pdf(self):
        r = self._window.create_file_dialog(webview.OPEN_DIALOG, directory=DESKTOP,
                                           file_types=("PDF files (*.pdf)",))
        return r[0] if r else None

    # ---- conversion ------------------------------------------------------
    def convert(self, pdf_path, invoice_no):
        if self._busy:
            return {"error": "Already converting a file."}
        invoice_no = (invoice_no or "").strip()
        if not invoice_no:
            return {"error": "Enter an invoice number first."}
        if not pdf_path or not os.path.exists(pdf_path):
            return {"error": "That PDF can't be found anymore. Drop it in again."}
        tess = conv.find_tesseract()
        if not tess:
            if IS_WIN:
                return {"error": "The text reader (Tesseract) is missing from this copy of the app. "
                                 "Download the app again."}
            return {"error": "Tesseract isn't installed on this Mac. Open Terminal, run "
                             "'brew install tesseract', then try again."}
        pytesseract.pytesseract.tesseract_cmd = tess

        s = load_settings()
        start = s.get("last_folder") if os.path.isdir(s.get("last_folder", "")) else DESKTOP
        folder = self._window.create_file_dialog(webview.FOLDER_DIALOG, directory=start)
        if not folder:
            return {"cancelled": True}
        folder = folder[0] if isinstance(folder, (list, tuple)) else folder

        self._busy = True
        try:
            def progress(p, n):
                self._window.evaluate_js(f"window.onProgress({p}, {n})")
            store = load_names()
            known = lab.UW_NAMES + store["known"] + list(store["fixes"].values())
            groups = conv.read_pdf(pdf_path, template_path(), progress=progress, known_names=known)
        except Exception as e:
            traceback.print_exc()
            return {"error": f"Couldn't read this PDF: {e}"}
        finally:
            self._busy = False

        # spelling fixes made in earlier weeks
        fixes = {lab._key(k): v for k, v in store["fixes"].items()}
        lab.apply_renames(groups, {e["name"]: fixes[lab._key(e["name"])]
                                   for g in groups for e in g.get("data", {}).get("employees", [])
                                   if lab._key(e["name"]) in fixes})
        review = lab.names_to_review(groups, known)

        people = lab.label_people(groups)
        checklist = load_checks()
        matches, missing = lab.match_checks(people, checklist)
        self._pending = {"groups": groups, "folder": folder, "invoice_no": invoice_no,
                         "checklist": checklist}
        if not matches and not review:
            return self.finish([], [])
        matched = {m["name"] for m in matches}
        return {"checks": {
            "names": review,
            "rows": matches,
            "missing": [e["name"] for e in checklist if e["name"] in missing and e.get("amount")],
            "others": [{"name": p["name"], "total": lab.dollars(p["total"])}
                       for p in people if p["name"] not in matched],
        }}

    def cancel_pending(self):
        self._pending = None
        return True

    def finish(self, confirmed, names=None):
        """Write the invoices and the label sheet, using what the user confirmed
        in the pop-up: check amounts [{name, amount, entry, remove}] and names
        [{from, to}] (to = the corrected spelling, or the same name)."""
        pend = self._pending
        if not pend:
            return {"error": "Nothing to save. Drop the PDF in again."}
        confirmed = confirmed or []
        names = names or []
        folder = pend["folder"]
        try:
            # names first: the invoice, the labels and the check list all use them
            renames = {}
            for n in names:
                old, new = (n.get("from") or "").strip(), " ".join((n.get("to") or "").split())
                if old and new and new != old:
                    renames[old] = new
            lab.apply_renames(pend["groups"], renames)
            if names:
                store = load_names()
                keys = {lab._key(k) for k in store["known"]}
                for n in names:
                    final = renames.get(n["from"], n["from"])
                    if lab._key(final) not in keys:
                        store["known"].append(final)
                        keys.add(lab._key(final))
                for old, new in renames.items():
                    store["fixes"][old] = new
                save_names(store)
            people = lab.label_people(pend["groups"])

            for c in confirmed:
                c["name"] = renames.get(c["name"], c["name"])
                c["amount"] = round(float(c.get("amount") or 0), 2)
            if confirmed:
                save_checks(lab.update_checklist(pend["checklist"], people, confirmed))
            checks = {c["name"]: c["amount"] for c in confirmed if not c.get("remove") and c["amount"]}
            results = conv.write_invoices(pend["groups"], template_path(), pend["invoice_no"], folder,
                                          overwrite=False)
            done = [r for r in results if "error" not in r]
            labels_out = None
            if people:
                first = done[0]["invoice_no"] if done else pend["invoice_no"]
                path = conv.unique_path(os.path.join(folder, f"MLL_LABELS_{first}.docx"))
                lab.build_labels(labels_template_path(), path, people, checks)
                labels_out = {"path": path, "filename": os.path.basename(path),
                              "count": len(people), "checks": len(checks)}
        except Exception as e:
            traceback.print_exc()
            return {"error": f"Couldn't save: {e}"}
        self._pending = None

        s = load_settings()
        if done:
            s["last_invoice_no"] = done[-1]["invoice_no"]
        s["last_folder"] = folder
        save_settings(s)

        out = []
        for r in results:
            item = {"page": r["page"], "label": r["label"]}
            if "error" in r:
                item["error"] = r["error"]
            else:
                ws, we = r["week_start"], r["week_end"]
                item.update(
                    invoice_no=r["invoice_no"], employees=r["employees"],
                    week=f"{ws:%b} {ws.day}" + (f"–{we.day}" if ws.month == we.month else f" – {we:%b} {we.day}") + f", {we.year}",
                    total=f"${r['total']:,.2f}",
                    pdf_total=f"${r['pdf_total']:,.2f}" if r["pdf_total"] is not None else None,
                    total_ok=r["total_ok"], issues=r["issues"], path=r["path"],
                    filename=os.path.basename(r["path"]))
            out.append(item)
        return {"results": out, "folder": folder, "labels": labels_out,
                "next_invoice_no": conv.next_invoice_no(s["last_invoice_no"], 1) if done else pend["invoice_no"]}

    # ---- Finder / Explorer / Excel ----------------------------------------
    def open_file(self, path):
        if IS_WIN:
            os.startfile(path)  # noqa: opens in Excel (or the default app for .xlsx)
        else:
            subprocess.run(["open", path])

    def reveal(self, path):
        if IS_WIN:
            subprocess.run(["explorer", "/select,", os.path.normpath(path)])
        else:
            subprocess.run(["open", "-R", path])


def bind_drop(window):
    def on_drop(e):
        files = e.get("dataTransfer", {}).get("files", [])
        paths = [f.get("pywebviewFullPath") for f in files if f.get("pywebviewFullPath")]
        pdfs = [p for p in paths if p.lower().endswith(".pdf")]
        if pdfs:
            window.evaluate_js(f"window.onFileDropped({json.dumps(pdfs[0])})")
        elif paths:
            window.evaluate_js("window.onWrongFile()")

    def noop(e):
        pass

    try:
        window.events.loaded.wait(15)
    except Exception:
        pass
    doc = window.dom.document
    doc.events.dragenter += DOMEventHandler(noop, True, True)
    doc.events.dragover += DOMEventHandler(noop, True, True, debounce=500)
    doc.events.drop += DOMEventHandler(on_drop, True, True)


def main():
    api = Api()
    # Load the page as text rather than a file path: inside a built .app the
    # web view may not be allowed to read files from the bundle (blank window).
    with open(os.path.join(RES_DIR, "ui.html"), encoding="utf-8") as f:
        html = f.read()
    window = webview.create_window(
        APP_NAME, html=html, js_api=api,
        width=560, height=660, min_size=(460, 520))
    api._window = window  # underscore: pywebview must not scan the window object
    # Set IC_DEBUG=1 to get a right-click "Inspect Element" for troubleshooting.
    webview.start(bind_drop, window, debug=bool(os.environ.get("IC_DEBUG")))


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise
