import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest.mock import patch
import extraer_licitaciones as app


def entry(status='PUB', cpv='45210000', location='<CountrySubentityCode>ES521</CountrySubentityCode>'):
    return ET.fromstring(f'''<entry xmlns="http://www.w3.org/2005/Atom"><id>id-1</id>
    <ContractFolderStatus><ContractFolderStatusCode>{status}</ContractFolderStatusCode>
    <ContractFolderID>TEST/01</ContractFolderID><ProcurementProject>
    <RequiredCommodityClassification><ItemClassificationCode>{cpv}</ItemClassificationCode></RequiredCommodityClassification>
    <RealizedLocation>{location}</RealizedLocation></ProcurementProject></ContractFolderStatus></entry>''')


class ExtractorTests(unittest.TestCase):
    def test_missing_fields_remain_empty(self):
        row = app.extract(entry(), [])
        self.assertEqual(row['expediente'], 'TEST/01')
        self.assertEqual(row['presupuesto base sin IVA'], '')
        self.assertEqual(row['URL'], '')

    def test_filters_status_and_cpv(self):
        self.assertIsNone(app.extract(entry(status='ADJ'), []))
        self.assertIsNone(app.extract(entry(cpv='72000000'), []))

    def test_generic_neighbor_province_is_not_municipality(self):
        self.assertIsNone(app.extract(entry(location='<CountrySubentity>Región de Murcia</CountrySubentity>'), ['murcia']))
        self.assertIsNotNone(app.extract(entry(location='<Address><CityName>Murcia</CityName></Address>'), ['murcia']))

    def test_next_links_and_ten_page_limit(self):
        calls = []
        def download(url):
            calls.append(url)
            return f'<feed xmlns="http://www.w3.org/2005/Atom"><link rel="next" href="/page/{len(calls)}"/>{ET.tostring(entry(), encoding="unicode")}</feed>'.encode()
        with tempfile.TemporaryDirectory() as directory, patch.object(app, 'download', download), patch('builtins.print'):
            target = Path(directory)
            app.main(['--salida', str(target/'out.csv'), '--directorio-cache', str(target/'cache')])
            self.assertEqual(len(calls), 10)
            import json
            report = json.loads((target/'out.resumen.json').read_text(encoding='utf-8'))
            self.assertEqual(report['entradas_leidas'], 10)
            self.assertEqual(report['licitaciones'], 1)
            self.assertTrue(calls[1].endswith('/page/1'))


if __name__ == '__main__':
    unittest.main()
