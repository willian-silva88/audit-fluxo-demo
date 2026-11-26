import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app'))
from demo import Demo, ROOT


class TestDemo(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.demo = Demo(temp.name)

    def confirmar(self, item, ids, justificativa=''):
        return self.demo.review({
            'id': item['id'], 'selected_match_ids': ids,
            'justification': justificativa, 'review_confirmed': True,
        })

    def test_amostra_e_reimportacao(self):
        self.assertEqual(self.demo.engine.state()['requests'], [])
        estado = self.demo.load_sample()
        self.assertEqual(estado['stats']['payments'], 7)
        self.assertEqual([item['match_status'] for item in estado['requests']], [
            'encontrado', 'conferir', 'conferir', 'nao_encontrado', 'nao_encontrado', 'encontrado',
        ])
        item = self.demo.engine.request_detail(estado['requests'][0]['id'])
        self.confirmar(item, [item['matches'][0]['id']])
        self.assertEqual(self.demo.load_sample()['stats']['reviewed'], 1)

    def test_excecao_exige_justificativa_e_aparece_na_planilha(self):
        item = self.demo.engine.request_detail(self.demo.load_sample()['requests'][1]['id'])
        self.assertEqual(item['matched_amount_cents'], 185000)
        self.assertEqual(item['selected_amount_cents'], 0)
        ids = [item['matches'][0]['id']]
        with self.assertRaisesRegex(ValueError, 'justificativa'):
            self.confirmar(item, ids)
        revisado = self.confirmar(item, ids, 'Decisão de ensaio.')
        self.assertEqual(revisado['review_status'], 'conferido')
        self.assertEqual(revisado['selected_amount_cents'], 120000)
        self.assertEqual(revisado['history'][0]['event'], 'review_updated')
        with ZipFile(BytesIO(self.demo.workbook())) as arquivo:
            planilhas = ET.fromstring(arquivo.read('xl/workbook.xml'))
            self.assertEqual([aba.attrib['name'] for aba in planilhas.findall('.//{*}sheet')],
                             ['Resumo', 'Detalhamento'])
            resumo = ET.fromstring(arquivo.read('xl/worksheets/sheet1.xml'))
            self.assertIn('Decisão de ensaio.', [node.text for node in resumo.findall('.//{*}t')])
            self.assertFalse(any('externalLinks' in nome for nome in arquivo.namelist()))

    def test_filial_ausente_e_tipo_fora_do_criterio(self):
        itens = self.demo.load_sample()['requests']
        ambiguo = self.demo.engine.request_detail(itens[2]['id'])
        self.assertEqual({baixa['branch'] for baixa in ambiguo['matches']}, {'DEMO-SUL', 'DEMO-NORTE'})
        with self.assertRaisesRegex(ValueError, 'filial'):
            self.confirmar(ambiguo, [ambiguo['matches'][0]['id']], 'Ensaio.')
        resolvido = self.demo.review({'id': ambiguo['id'], 'branch': 'DEMO-SUL'})
        self.assertEqual([baixa['branch'] for baixa in resolvido['matches']], ['DEMO-SUL'])
        fora = self.demo.engine.request_detail(itens[3]['id'])
        self.assertFalse(fora['matches'][0]['eligible'])
        with self.assertRaisesRegex(ValueError, 'tipo configurado'):
            self.confirmar(fora, [fora['matches'][0]['id']], 'Ensaio.')

    def test_troca_de_base_invalida_conferencia_afetada(self):
        item = self.demo.engine.request_detail(self.demo.load_sample()['requests'][0]['id'])
        self.confirmar(item, [item['matches'][0]['id']])
        caminho = Path(self.demo.directory) / 'nova-base-sintetica.csv'
        origem = (ROOT / 'exemplos/baixas-sinteticas.csv').read_text(encoding='utf-8')
        caminho.write_text(origem.replace('1240,00', '1245,00'), encoding='utf-8')
        self.demo.engine.import_payments(caminho)
        atualizado = self.demo.engine.request_detail(item['id'])
        self.assertTrue(atualizado['review_invalidated'])
        self.assertEqual(atualizado['review_status'], 'em_conferencia')
        self.assertEqual(atualizado['history'][0]['event'], 'review_invalidated')


if __name__ == '__main__':
    unittest.main()
