import re
import sys
import openpyxl


def este_pagina_de_orar(nume_pagina):
    # Paginile cu orarul propriu-zis se numesc "AN 1", "AN 2", "AN3" ... (nu "Anunturi - AN1")
    return re.fullmatch(r'\s*AN\s*0?\d\s*', nume_pagina, re.IGNORECASE) is not None


def curata_excel_orar(fisier_intrare, fisier_iesire):
    workbench = openpyxl.load_workbook(fisier_intrare)

    # Alegem paginile de orar dupa nume, nu dupa pozitie (daca secretariatul adauga o pagina noua, nu se strica nimic).
    # Daca nu gasim nicio pagina cu acest format, folosim vechea varianta: primele 4 pagini.
    pagini_orar = [p for p in workbench.sheetnames if este_pagina_de_orar(p)] or workbench.sheetnames[:4]

    # Acum parcurgem fiecare pagina pe rand
    for nume_pagina in pagini_orar:
        sheet = workbench[nume_pagina]

        merged_ranges = list(sheet.merged_cells.ranges)

        for merged_range in merged_ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            top_left_cell_value = sheet.cell(row=min_row, column=min_col).value

            sheet.unmerge_cells(str(merged_range))

            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    sheet.cell(row=row, column=col).value = top_left_cell_value

    print(f"Salvez varianta curatata in {fisier_iesire}")
    workbench.save(fisier_iesire)
    return fisier_iesire


if __name__ == "__main__":
    # Folosire: python parser_orar.py [orar_etti.xlsx] [orar_curatat.xlsx]
    intrare = sys.argv[1] if len(sys.argv) > 1 else 'orar_etti.xlsx'
    iesire = sys.argv[2] if len(sys.argv) > 2 else 'orar_curatat.xlsx'
    curata_excel_orar(intrare, iesire)
