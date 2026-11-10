"""Checks the automated test conversions done by the built Windows app."""
import re, sys

EXPECTED = {  # PDF -> list of (employees, invoice total) per invoice
    "doc8.pdf": [(39, 31724.53), (47, 45771.43)],
    "combined.pdf": [(91, 92906.17)],
}
text = open(sys.argv[1], encoding="utf-8").read()
print(text)
ok = True
for pdf, exp in EXPECTED.items():
    block = text.split(f"=== {pdf}")[1].split("===")[0] if f"=== {pdf}" in text else ""
    got = re.findall(r"RESULT page=\S+ .*?employees=(\d+) total=([\d.]+) .*?total_ok=(\w+)", block)
    errors = re.findall(r"RESULT .* ERROR (.*)", block)
    if errors or len(got) != len(exp):
        print(f"FAIL {pdf}: expected {len(exp)} invoices, got {len(got)}; errors: {errors}"); ok = False; continue
    for (emp, tot, tok), (e_emp, e_tot) in zip(got, exp):
        if int(emp) != e_emp or abs(float(tot) - e_tot) > 0.005 or tok != "True":
            print(f"FAIL {pdf}: got {emp} employees ${tot} (total_ok={tok}), expected {e_emp} / ${e_tot}"); ok = False
        else:
            print(f"PASS {pdf}: {emp} employees, ${tot}, matches the PDF")
LABELS = {"combined.pdf": (91, 13)}  # PDF -> (labels, labels with a check)
for pdf, (n, k) in LABELS.items():
    block = text.split(f"=== {pdf}")[1].split("===")[0] if f"=== {pdf}" in text else ""
    m = re.search(r"LABELS people=(\d+) checks=(\d+)", block)
    if not m or (int(m[1]), int(m[2])) != (n, k):
        print(f"FAIL {pdf} labels: expected {n} labels / {k} checks, got {m.group(0) if m else 'nothing'}"); ok = False
    else:
        print(f"PASS {pdf} labels: {n} labels, {k} with a check")
sys.exit(0 if ok else 1)
