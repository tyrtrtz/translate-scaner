"""Update metadata, corrupt downloads, and the real Windows replacement script."""

import hashlib
import io
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError

from updater import ASSET_NAME, REPOSITORY, check_for_update, download_update, start_installation


class Updates(unittest.TestCase):
    def test_release_and_verified_download(self):
        executable = bytearray(128)
        executable[:2] = b'MZ'
        struct.pack_into('<I', executable, 60, 64)
        executable[64:70] = b'PE\x00\x00\x64\x86'
        url = f'https://github.com/{REPOSITORY}/releases/download/v1.2.0/{ASSET_NAME}'
        asset = dict(name=ASSET_NAME, browser_download_url=url, state='uploaded', size=len(executable),
                     digest='sha256:' + hashlib.sha256(executable).hexdigest())
        release = dict(tag_name='v1.2.0', draft=False, prerelease=False, assets=[asset])

        def check(current='1.1.0'):
            with patch('updater.urlopen', return_value=io.BytesIO(json.dumps(release).encode())):
                return check_for_update(current)

        result = check()
        self.assertEqual(result['version'], '1.2.0')
        self.assertIsNone(check('1.2.0'))
        self.assertIsNone(check('1.10.0'))
        release['prerelease'] = True
        self.assertIsNone(check())
        release['prerelease'] = False
        for key, invalid in (('browser_download_url', 'https://example.com/app.exe'), ('digest', None), ('size', -1)):
            original = asset[key]
            asset[key] = invalid
            with self.assertRaises(ValueError):
                check()
            asset[key] = original
        with patch('updater.urlopen', side_effect=HTTPError(url, 404, 'Not found', {}, None)):
            self.assertIsNone(check_for_update())
        with patch('updater.urlopen', side_effect=HTTPError(url, 403, 'Rate limited', {}, None)):
            with self.assertRaises(HTTPError):
                check_for_update()

        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "客户's [程序].exe"
            target.write_bytes(b'old executable')
            percentages = []
            with patch('updater.urlopen', return_value=io.BytesIO(executable)):
                download = download_update(result, target, percentages.append)
            self.assertEqual(download.read_bytes(), executable)
            self.assertEqual(percentages[-1], 100)
            shutil.rmtree(download.parent)
            for content in (executable[:-1], executable + b'extra', b'X' * len(executable)):
                with patch('updater.urlopen', return_value=io.BytesIO(content)), self.assertRaises(ValueError):
                    download_update(result, target)
                self.assertEqual(target.read_bytes(), b'old executable')
                self.assertEqual(list(Path(directory).iterdir()), [target])
            invalid = b'X' * len(executable)
            with patch('updater.urlopen', return_value=io.BytesIO(invalid)), self.assertRaises(ValueError):
                download_update(dict(result, sha256=hashlib.sha256(invalid).hexdigest()), target)
            self.assertEqual(target.read_bytes(), b'old executable')

    @unittest.skipUnless(sys.platform == 'win32', 'Executable replacement uses Windows PowerShell')
    def test_windows_replacement_waits_and_handles_unicode_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory) / "客户's [更新]"
            base.mkdir()
            target = base / "程序.exe"
            staging = base / 'staging'
            staging.mkdir()
            download = staging / ASSET_NAME
            # Native console utilities exit immediately when run without args;
            # this exercises replacement/restart without leaving another GUI.
            shutil.copy2(Path(os.environ['SystemRoot']) / 'System32' / 'cmd.exe', target)
            shutil.copy2(Path(os.environ['SystemRoot']) / 'System32' / 'where.exe', download)
            expected = hashlib.sha256(download.read_bytes()).digest()
            old = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(1)'])
            self.addCleanup(lambda: old.poll() is None and old.kill())
            helper = start_installation(download, target, old.pid)
            try:
                self.assertEqual(helper.wait(timeout=30), 0)
            finally:
                if helper.poll() is None:
                    helper.kill()
            self.assertIsNotNone(old.poll())
            self.assertEqual(hashlib.sha256(target.read_bytes()).digest(), expected)
            self.assertFalse(staging.exists())
            # A replacement failure after backing up must restore the old exe.
            # Suppress only the modal failure dialog in the unattended CI run.
            staging.mkdir()
            real_popen = subprocess.Popen

            def without_dialog(command, **kwargs):
                script = Path(command[-1])
                source = script.read_text(encoding='utf-8-sig')
                source = re.sub(r'    \[System.Windows.Forms.MessageBox\].*', '    Write-Output $failure', source)
                script.write_text(source, encoding='utf-8-sig')
                return real_popen(command, **kwargs)

            with patch('updater.subprocess.Popen', side_effect=without_dialog):
                helper = start_installation(download, target, old.pid)
            try:
                self.assertEqual(helper.wait(timeout=30), 1)
            finally:
                if helper.poll() is None:
                    helper.kill()
            self.assertEqual(hashlib.sha256(target.read_bytes()).digest(), expected)
            self.assertTrue((staging / 'update-error.txt').exists())


if __name__ == '__main__':
    unittest.main()
