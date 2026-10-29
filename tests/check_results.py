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
sys.exit(0 if ok else 1)
