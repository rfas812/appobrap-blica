"""Extrae hasta diez páginas ATOM de PLACSP. Solo requiere Python estándar.

El ámbito adicional se configura con --adyacentes (nombres separados por ;).
No se infiere el lugar de ejecución a partir del órgano de contratación.
PUB se interpreta como el estado publicado por la plataforma, sin deducir
otro estado a partir de la fecha límite.
"""
import argparse
import csv
import json
from datetime import datetime, timezone
import re
import time
import unicodedata
import urllib.request
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import urljoin

URL = 'https://contrataciondelestado.es/sindicacion/sindicacion_643/licitacionesPerfilesContratanteCompleto3.atom'
ATOM = '{http://www.w3.org/2005/Atom}'
ADYACENTES = ('Abanilla;Jumilla;Yecla;Santomera;Beniel;Murcia;San Javier;San Pedro del Pinatar;'
              'Almansa;Caudete;La Font de la Figuera;Fuente la Higuera;Fontanars dels Alforins;'
              'Ontinyent;Onteniente;Bocairent;Bocairente;Agullent;Benissoda;Benisoda;Albaida;Atzeneta d\'Albaida;'
              'Adzaneta de Albaida;Bèlgida;Bélgida;Bufali;El Palomar;Otos;Beniatjar;Salem;Aielo de Rugat;Ayelo de Rugat;Castelló de Rugat;Castellón de Rugat;'
              'Montitxelvo;Montichelvo;Terrateig;Vilallonga;Villalonga;Oliva')
FIELDS = ['expediente', 'órgano de contratación', 'objeto', 'CPV',
          'presupuesto base sin IVA', 'procedimiento',
          'fecha límite de presentación', 'lugar de ejecución', 'URL']

def normalize(value):
    return ''.join(c for c in unicodedata.normalize('NFKD', value.casefold())
                   if not unicodedata.combining(c))

def local(element):
    return element.tag.rsplit('}', 1)[-1]

def child(element, name):
    return next((e for e in element if local(e) == name), None) if element is not None else None

def path(element, *names):
    for name in names:
        element = child(element, name)
    return element

def value(element, *names):
    e = path(element, *names)
    return (e.text or '').strip() if e is not None else ''

def amount(row):
    try:
        return Decimal(row['presupuesto base sin IVA'])
    except InvalidOperation:
        return Decimal(0)

def extract(entry, adjacent):
    status = child(entry, 'ContractFolderStatus')
    if value(status, 'ContractFolderStatusCode') != 'PUB':
        return None
    project = child(status, 'ProcurementProject')
    if project is None:
        return None
    cpvs = list(dict.fromkeys((e.text or '').strip() for e in project.iter()
                             if local(e) == 'ItemClassificationCode' and (e.text or '').strip()))
    # Incluye CPV del proyecto o de sus lotes, sin sumar presupuestos de lotes.
    for lot in status:
        if local(lot) == 'ProcurementProjectLot':
            cpvs.extend((e.text or '').strip() for e in lot.iter()
                        if local(e) == 'ItemClassificationCode' and (e.text or '').strip())
    cpvs = list(dict.fromkeys(cpvs))
    if not any(c.startswith('45') for c in cpvs):
        return None
    locations = [e for e in project.iter() if local(e) == 'RealizedLocation']
    locations += [e for lot in status if local(lot) == 'ProcurementProjectLot'
                  for e in lot.iter() if local(e) == 'RealizedLocation']
    accepted = False
    places = []
    for loc in locations:
        parts = [(e.text or '').strip() for e in loc.iter()
                 if local(e) in {'CountrySubentity', 'CountrySubentityCode', 'CityName', 'PostalZone', 'Description', 'Line'}
                 and (e.text or '').strip()]
        places.append(', '.join(dict.fromkeys(parts)))
        for e in loc.iter():
            text = (e.text or '').strip()
            norm = normalize(text)
            tag = local(e)
            if tag == 'CountrySubentityCode' and text == 'ES521':
                accepted = True
            elif tag in {'CountrySubentity', 'CityName', 'Description', 'Line'}:
                if re.search(r'\b(alicante|alacant)\b', norm) or (tag != 'CountrySubentity' and any(
                    re.search(r'(?<!\w)' + re.escape(m) + r'(?!\w)', norm) for m in adjacent)):
                    accepted = True
            elif tag == 'PostalZone' and re.fullmatch(r'03\d{3}', text):
                accepted = True
    if not accepted:
        return None
    process = child(status, 'TenderingProcess')
    deadline = child(process, 'TenderSubmissionDeadlinePeriod')
    date = value(deadline, 'EndDate')
    clock = value(deadline, 'EndTime')
    links = entry.findall(ATOM + 'link')
    link = next((e.get('href', '') for e in links if e.get('rel', 'alternate') == 'alternate'), '')
    return dict(zip(FIELDS, [
        value(status, 'ContractFolderID'),
        value(status, 'LocatedContractingParty', 'Party', 'PartyName', 'Name'),
        value(project, 'Name'), '; '.join(cpvs),
        value(project, 'BudgetAmount', 'TaxExclusiveAmount'),
        value(process, 'ProcedureCode'),
        ' '.join(x for x in (date, clock) if x),
        '; '.join(dict.fromkeys(p for p in places if p)), link]))

def download(url):
    for attempt in range(3):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Python PLACSP feed reader'})
            with urllib.request.urlopen(req, timeout=120) as response:
                return response.read()
        except Exception:
            if attempt == 2:
                raise
            time.sleep(2 * (attempt + 1))

def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--adyacentes', default=ADYACENTES)
    parser.add_argument('--primera-pagina', type=Path)
    parser.add_argument('--cache', action='store_true', help='Reprocesa las páginas ya descargadas')
    parser.add_argument('--directorio-cache', type=Path, default=Path('feed_paginas'))
    parser.add_argument('--max-paginas', type=int, choices=range(1, 11), default=10)
    parser.add_argument('--salida', type=Path, default=Path('resultados/licitaciones_obras.csv'))
    parser.add_argument('--excel', action='store_true', help='Genera también un Excel (requiere openpyxl)')
    args = parser.parse_args(argv)
    if args.excel:
        try:
            from exportar_excel import export_excel
        except ImportError as exc:
            parser.error(f'No se puede cargar el exportador Excel: {exc}. Instala requirements.txt.')
    adjacent = [normalize(m.strip()) for m in args.adyacentes.split(';') if m.strip()]
    url, visited, rows, entries, pages = URL, set(), {}, 0, []
    for index in range(args.max_paginas):
        if not url or url in visited:
            break
        visited.add(url)
        cache = args.directorio_cache
        cached = cache / f'pagina_{index + 1:02}.atom'
        data = cached.read_bytes() if args.cache else (args.primera_pagina.read_bytes() if index == 0 and args.primera_pagina else download(url))
        cache.mkdir(parents=True, exist_ok=True)
        (cache / f'pagina_{index + 1:02}.atom').write_bytes(data)
        root = ET.fromstring(data)
        batch = root.findall(ATOM + 'entry')
        entries += len(batch)
        pages.append({'url': url, 'entradas': len(batch)})
        for entry in batch:
            row = extract(entry, adjacent)
            # El feed puede repetir una licitación; prevalece su primera aparición.
            key = value(entry, 'id') or ('sin-id', index, len(rows), ET.tostring(entry))
            if key not in rows:
                rows[key] = row
        print(f'Página {index + 1}: {len(batch)} entradas', flush=True)
        next_link = next((e.get('href') for e in root.findall(ATOM + 'link') if e.get('rel') == 'next'), None)
        url = urljoin(url, next_link) if next_link else None
    selected = sorted((r for r in rows.values() if r), key=amount, reverse=True)
    args.salida.parent.mkdir(parents=True, exist_ok=True)
    with args.salida.open('w', encoding='utf-8-sig', newline='') as output:
        writer = csv.DictWriter(output, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(selected)
    total = sum((amount(r) for r in selected), Decimal(0))
    report = {'fecha_extraccion_UTC': datetime.now(timezone.utc).isoformat(),
              'paginas': pages, 'entradas_leidas': entries, 'licitaciones': len(selected),
              'importe_total_sin_IVA': str(total), 'adyacentes': args.adyacentes,
              'presupuestos_ausentes': sum(not r['presupuesto base sin IVA'] for r in selected),
              'top_5': selected[:5]}
    args.salida.with_suffix('.resumen.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    if args.excel:
        export_excel(selected, report, args.salida.with_suffix('.xlsx'))
    print(f'Entradas leídas en total: {entries}\nLicitaciones: {len(selected)}\nImporte total sin IVA: {total:,.2f} EUR')
    print('Las 5 de mayor importe:')
    for row in selected[:5]:
        print(f"{amount(row):,.2f} EUR | {row['expediente']} | {row['objeto']} | {row['órgano de contratación']}")

if __name__ == '__main__':
    main()
