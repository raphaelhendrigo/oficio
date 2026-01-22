import importlib.util
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

selectors_path = ROOT / "src" / "selectors.py"
spec = importlib.util.spec_from_file_location("selectors", selectors_path)
module = importlib.util.module_from_spec(spec) if spec else None
if spec and spec.loader and module:
    spec.loader.exec_module(module)
    sys.modules["selectors"] = module

from bot import read_yaml_steps


class TestSteps(unittest.TestCase):
    def test_steps_yaml_loads(self) -> None:
        steps = read_yaml_steps(Path("docs/steps.yaml"))
        self.assertTrue(steps)

    def test_actions_supported(self) -> None:
        steps = read_yaml_steps(Path("docs/steps.yaml"))
        actions = {str(s.get("action", "")).strip().lower() for s in steps}
        supported = {
            "goto",
            "click",
            "type",
            "select",
            "wait",
            "assert",
            "download",
            "upload",
            "select_row",
            "click_row_icon",
            "close_page",
            "filter_grid",
        }
        self.assertTrue(actions.issubset(supported))
