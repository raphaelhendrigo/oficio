import os
import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from config import load_config


class TestConfig(unittest.TestCase):
    def test_env_overrides(self) -> None:
        env = {
            "ETCM_USER": "user@example",
            "ETCM_PASS": "pass",
            "VISOES": "Aposentadoria",
            "DISTRIBUIDO_PARA": "silvana",
            "PROCESSOS_LIST": "TC/001,TC/002",
        }
        with mock.patch.dict(os.environ, env, clear=False):
            cfg = load_config()
        self.assertEqual(cfg.etcm_user, "user@example")
        self.assertEqual(cfg.etcm_pass, "pass")
        self.assertEqual(cfg.visoes, "Aposentadoria")
        self.assertEqual(cfg.distribuido_para, "silvana")
        self.assertEqual(cfg.process_list, ["TC/001", "TC/002"])
