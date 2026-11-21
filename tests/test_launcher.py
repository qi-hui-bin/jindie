import os
import re
import subprocess
import sys
import time
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.request import urlopen

from .voucher_fixtures import write_example


@unittest.skipUnless(sys.platform == "win32", "Windows desktop launcher")
class DesktopLauncherTests(unittest.TestCase):
    def test_launcher_opens_loopback_and_keeps_demo_separate(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            bundle = root / "demo release"
            write_example(bundle / "demo")
            local = root / "profile"
            env = {**os.environ, "LOCALAPPDATA": str(local)}
            command = [sys.executable, "-X", "utf8", "-m", "jindie.launcher", "--bundle-root", str(bundle), "--no-browser"]
            for demo in (True, False):
                args = [*command, "--demo"] if demo else command
                log_path = root / "launcher.log"
                with log_path.open("w+b") as log:
                    process = subprocess.Popen(args, env=env, stdout=log, stderr=subprocess.STDOUT)
                    try:
                        deadline = time.monotonic() + 15
                        address = None
                        while time.monotonic() < deadline:
                            match = re.search(r"http://127\.0\.0\.1:[0-9]+/", log_path.read_text(encoding="utf-8"))
                            if match:
                                address = match[0]
                                break
                            if process.poll() is not None:
                                self.fail(log_path.read_text(encoding="utf-8"))
                            time.sleep(0.05)
                        if address is None:
                            self.fail(log_path.read_text(encoding="utf-8"))
                        with urlopen(address, timeout=5) as response:
                            self.assertIn("合成演示账套" if demo else "新增账套", response.read().decode("utf-8"))
                        with urlopen(address + "configuration", timeout=5) as response:
                            page = response.read().decode("utf-8")
                            if demo:
                                self.assertNotIn("新增账套", page)
                            else:
                                self.assertIn("新增账套", page)
                        self.assertTrue((local / "Jindie" / "demo" / "statements.sqlite3").is_file())
                        self.assertEqual((local / "Jindie" / "statements.sqlite3").exists(), not demo)
                        duplicate = subprocess.run(args, env=env, capture_output=True, timeout=10)
                        self.assertEqual(duplicate.returncode, 2)
                        self.assertIn("已启动", duplicate.stderr.decode("utf-8"))
                    finally:
                        process.terminate()
                        process.wait(10)
