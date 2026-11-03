import base64
import json
import os
import shutil
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook
from PIL import Image as PILImage

from qc_sheet.generator import generate_qc_sheet, StyleNotFound, SizeNotFound, NoMeasurements
from qc_sheet.storage import GitHubStorage


class GeneratorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        self.spec, self.template, self.logo = root / 'spec.xlsx', root / 'template.xlsx', root / 'logo.png'
        book = Workbook()
        sheet = book.active
        sheet.append(['STYLE-001'])
        sheet.append(['', '', 'M', 'L'])
        sheet.append(['', 'Chest', 50, 55])
        sheet.append(['', '가슴', None, None])
        sheet.append(['', 'Length', 70, 75])
        sheet.append(['', '기장', None, None])
        book.save(self.spec)
        template = Workbook()
        template.active['H1'] = '원본 양식 유지'
        template.save(self.template)
        PILImage.new('RGB', (8, 8), 'black').save(self.logo)

    def generate(self, style='style-001', size='M', language='English', fallback=False):
        return generate_qc_sheet(self.spec, self.template, self.logo, style, size, language, fallback)

    def test_result_contains_values_formula_logo_and_template(self):
        book = load_workbook(BytesIO(self.generate()))
        sheet = book.active
        self.assertEqual((sheet['B6'].value, sheet['G6'].value), ('style-001', 'M'))
        self.assertEqual((sheet['A9'].value, sheet['B9'].value), ('Chest', 50))
        self.assertEqual((sheet['A10'].value, sheet['B10'].value), ('Length', 70))
        self.assertEqual(sheet['D9'].value, '=IF(C9="","",IFERROR(C9-B9,""))')
        self.assertEqual(sheet['H1'].value, '원본 양식 유지')
        self.assertEqual(len(sheet._images), 1)

    def test_korean_measurements_and_fallback(self):
        with self.assertRaises(StyleNotFound):
            self.generate(style='NOT-FOUND')
        book = load_workbook(BytesIO(self.generate(style='NOT-FOUND', language='Korean', fallback=True)))
        self.assertEqual((book.active['A9'].value, book.active['B9'].value), ('가슴', 50))
        self.assertEqual((book.active['A10'].value, book.active['B10'].value), ('기장', 70))

    def test_missing_size_and_measurements(self):
        with self.assertRaises(SizeNotFound):
            self.generate(size='XXL')
        book = load_workbook(self.spec)
        book.active['B3'] = '123'
        book.active['B4'] = '123'
        book.active['B5'] = '123'
        book.active['B6'] = '123'
        book.save(self.spec)
        with self.assertRaises(NoMeasurements):
            self.generate()


class StorageTests(unittest.TestCase):
    def setUp(self):
        self.storage = GitHubStorage('test-token', 'example/test', 'main')

    @patch('qc_sheet.storage.requests.put')
    @patch('qc_sheet.storage.requests.get')
    def test_upload_preserves_path_existing_sha_and_contents(self, get, put):
        get.return_value = Mock(status_code=200, json=lambda: {'sha': 'old-sha'})
        put.return_value = Mock(status_code=200)
        with tempfile.TemporaryDirectory() as directory:
            file = Path(directory) / 'sample.xlsx'
            file.write_bytes(b'example')
            self.storage.upload(file, 'spec/sample.xlsx')
        payload = json.loads(put.call_args.kwargs['data'])
        self.assertEqual(payload['sha'], 'old-sha')
        self.assertEqual(payload['branch'], 'main')
        self.assertEqual(base64.b64decode(payload['content']), b'example')
        self.assertTrue(put.call_args.args[0].endswith('/spec/sample.xlsx'))

    @patch('qc_sheet.storage.requests.delete')
    @patch('qc_sheet.storage.requests.get')
    def test_delete_uses_sha_and_failure_does_not_delete(self, get, delete):
        get.return_value = Mock(status_code=404)
        self.assertFalse(self.storage.delete('image/sample.png'))
        delete.assert_not_called()
        get.return_value = Mock(status_code=200, json=lambda: {'sha': 'old-sha'})
        delete.return_value = Mock(status_code=200)
        self.assertTrue(self.storage.delete('image/sample.png'))
        self.assertEqual(delete.call_args.kwargs['json']['sha'], 'old-sha')
        self.assertEqual(delete.call_args.kwargs['json']['branch'], 'main')

    @patch('qc_sheet.storage.subprocess.run')
    def test_sync_copies_same_folders(self, run):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            cache = root / 'repo_cache'
            (cache / 'spec').mkdir(parents=True)
            (cache / 'spec' / 'sample.xlsx').write_bytes(b'spec')
            (root / 'uploaded' / 'spec').mkdir(parents=True)
            with patch('qc_sheet.storage.REPO_LOCAL', cache), patch('qc_sheet.storage.BASE_DIR', str(root / 'uploaded')):
                self.storage.sync()
            self.assertEqual((root / 'uploaded' / 'spec' / 'sample.xlsx').read_bytes(), b'spec')
            self.assertIn('pull', run.call_args.args[0])


class ScreenTests(GeneratorTests):
    def test_screen_generation_and_first_sheet_confirmation(self):
        from qc_sheet import ui
        class Rerun(Exception):
            pass
        class FakeStreamlit:
            def __init__(self):
                self.secrets = {'GH_TOKEN': 'test-token', 'GH_REPO': 'example/test'}
                self.session_state = {}
                self.style = 'STYLE-001'
                self.clicked = ''
                self.successes, self.warnings, self.downloads = [], [], []
            def __enter__(self): return self
            def __exit__(self, *args): return False
            def file_uploader(self, *args, **kwargs): return None
            def columns(self, widths): return [self] * (widths if isinstance(widths, int) else len(widths))
            def expander(self, *args): return self
            def subheader(self, *args): pass
            def markdown(self, *args): pass
            def write(self, *args): pass
            def toast(self, *args, **kwargs): pass
            def error(self, message): raise AssertionError(message)
            def stop(self): raise AssertionError('Unexpected stop')
            def rerun(self): raise Rerun()
            def warning(self, message): self.warnings.append(message)
            def success(self, message): self.successes.append(message)
            def selectbox(self, label, options, **kwargs): return options[0]
            def text_input(self, label): return self.style if label == '스타일넘버 입력' else 'M'
            def button(self, label, **kwargs): return label == self.clicked
            def download_button(self, label, **kwargs): self.downloads.append(kwargs)
        st = FakeStreamlit()
        old_cwd = Path.cwd()
        try:
            os.chdir(self.temp.name)
            root = Path('uploaded')
            for folder in ('spec', 'template', 'image'):
                (root / folder).mkdir(parents=True)
            shutil.copy2(self.spec, root / 'spec' / 'spec.xlsx')
            shutil.copy2(self.template, root / 'template' / 'template.xlsx')
            shutil.copy2(self.logo, root / 'image' / 'logo.png')
            with patch.object(ui, 'st', st), patch.object(ui.GitHubStorage, 'sync'), patch.object(ui, 'github_file_date', return_value='날짜 확인 불가'):
                st.clicked = '🚀 QC시트 생성'
                ui.render_qc_sheet()
                self.assertTrue(any('생성 완료' in item for item in st.successes))
                self.assertTrue(any(item.get('file_name') == 'QC_STYLE-001_M.xlsx' for item in st.downloads))
                st.style = 'NOT-FOUND'
                with self.assertRaises(Rerun): ui.render_qc_sheet()
                st.clicked = '네, 사용합니다'
                ui.render_qc_sheet()
                self.assertTrue(any('맨 앞의 스펙' in item for item in st.warnings))
                self.assertTrue(any(item.get('file_name') == 'QC_NOT-FOUND_M.xlsx' for item in st.downloads))
        finally:
            os.chdir(old_cwd)


if __name__ == '__main__':
    unittest.main()
