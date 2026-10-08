"""Exportación portable a Excel; los datos de entrada se mantienen como texto literal."""
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

FIELDS = ['expediente', 'órgano de contratación', 'objeto', 'CPV',
          'presupuesto base sin IVA', 'procedimiento',
          'fecha límite de presentación', 'lugar de ejecución', 'URL']
CURRENCY = '#,##0.00" €"'


def export_excel(rows, report, destination):
    workbook = Workbook()
    summary = workbook.active
    summary.title = 'Resumen'
    detail = workbook.create_sheet('Licitaciones')
    sources = workbook.create_sheet('Fuentes y alcance')
    summary.append(['Licitaciones de obras: Alicante y municipios limítrofes'])
    summary.append(['Extracción (UTC)', report.get('fecha_extraccion_UTC', '')])
    summary.append(['Páginas leídas', len(report['paginas'])])
    summary.append(['Entradas leídas', report['entradas_leidas']])
    summary.append(['Licitaciones', len(rows)])
    summary.append(['Total sin IVA', float(Decimal(report['importe_total_sin_IVA']))])
    summary['B6'].number_format = CURRENCY
    summary.append([])
    summary.append(['Expediente', 'Objeto: cinco mayores importes', 'Presupuesto sin IVA'])
    for row in rows[:5]:
        summary.append([row['expediente'], row['objeto'],
                        float(Decimal(row[FIELDS[4]])) if row[FIELDS[4]] else None])
        summary.cell(summary.max_row, 3).number_format = CURRENCY
    detail.append(FIELDS + ['observación'])
    for row in rows:
        values = [row.get(field, '') or None for field in FIELDS]
        values[4] = float(Decimal(values[4])) if values[4] is not None else None
        if values[6]:
            try:
                values[6] = datetime.fromisoformat(values[6])
            except ValueError:
                pass  # Conserva el texto original si la fuente no usa fecha ISO.
        note = ('Revisar ubicación: ciudad Murcia y CP 30201.'
                if row['expediente'] == '2026/AR44U/00003772E' else None)
        detail.append(values + [note])
        detail.cell(detail.max_row, 5).number_format = CURRENCY
        detail.cell(detail.max_row, 7).number_format = 'dd/mm/yyyy hh:mm'
        if values[8]:
            detail.cell(detail.max_row, 9).hyperlink = values[8]
    if rows:
        table = Table(displayName='LicitacionesObras', ref=f'A1:J{detail.max_row}')
        table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        detail.add_table(table)
    detail.freeze_panes = 'B2'
    sources.append(['Criterio', 'Valor'])
    for pair in [('Estado', 'PUB según el feed; no se recalcula por fecha'),
                 ('CPV', 'Al menos un código con prefijo 45'),
                 ('Presupuesto', 'TaxExclusiveAmount; sin IVA; sin sumar lotes'),
                 ('Lugar', 'RealizedLocation del proyecto y lotes'),
                 ('Procedimiento', 'Código original del feed'),
                 ('Campos ausentes', 'Vacíos'),
                 ('Duplicados', 'Primera aparición de cada ID ATOM'),
                 ('Municipios adicionales', report['adyacentes'])]:
        sources.append(pair)
    sources.append([])
    sources.append(['Página', 'Entradas', 'URL'])
    for index, page in enumerate(report['paginas'], 1):
        sources.append([index, page['entradas'], page['url']])
    if any(row['expediente'] == '2026/AR44U/00003772E' for row in rows):
        summary.append([])
        summary.append(['Ubicación dudosa', '2026/AR44U/00003772E declara Murcia y CP 30201. Incluido por la ciudad del feed.'])
    for sheet, widths in [(summary, [36, 85, 25]),
                          (detail, [30, 48, 85, 26, 25, 20, 27, 45, 80, 48]),
                          (sources, [36, 100, 125])]:
        sheet.sheet_view.showGridLines = False
        for index, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        for line in sheet:
            for cell in line:
                cell.font = Font(name='Arial', size=10, color='243747')
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                if isinstance(cell.value, str):
                    cell.data_type = 's'  # Evita interpretar contenido del feed como fórmulas.
            sheet.row_dimensions[line[0].row].height = 75 if sheet == detail else 45
        for cell in sheet[1]:
            cell.fill = PatternFill('solid', fgColor='23435F')
            cell.font = Font(name='Arial', size=10, color='FFFFFF', bold=True)
    workbook.save(destination)
