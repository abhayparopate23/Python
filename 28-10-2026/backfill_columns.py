"""One-time helper: fill First Experience / Tenure (Overall) / Tenure (Current) / Remark
for profiles that are ALREADY in profiles.xlsx.

New captures fill these columns automatically, so you only need this for older rows.

How to run (keep app.py in the same folder):
    1. Close profiles.xlsx in Excel (and stop the service, to be safe).
    2. python backfill_columns.py
Rows whose four columns already contain anything are left alone (your manual edits are safe).
"""
import app


def main():
    with app.EXCEL_LOCK:
        wb, ws, cols = app.open_or_create_workbook()      # also adds the new columns if they are missing
        try:
            filled = skipped = empty = 0
            for r in range(2, ws.max_row + 1):
                exp = ws.cell(r, cols["experience section"]).value
                if not exp:
                    empty += 1
                    continue
                if any(ws.cell(r, cols[k]).value for k in app.DERIVED_KEYS):
                    skipped += 1
                    continue
                a = app.analyze_first_experience(str(exp))
                for key, field in (("first experience", "first_exp"), ("tenure overall", "overall"),
                                   ("tenure current", "current"), ("remark", "remark")):
                    cell = ws.cell(r, cols[key])
                    app.set_text(cell, a[field])
                    cell.alignment = app.WRAP
                filled += 1
            app.atomic_save_workbook(wb)
        finally:
            wb.close()
    print(f"Done. Filled {filled} row(s); {skipped} already had values; {empty} had no Experience text.")


if __name__ == "__main__":
    main()
