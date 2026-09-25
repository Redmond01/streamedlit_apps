from __future__ import annotations

from pathlib import Path
import sys
import traceback

# Ensure project root is in sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from tests import test_memory, test_gear, test_pos_recon, test_sales
import unittest

def run_tests():
    modules = [test_memory, test_gear, test_pos_recon]
    passed = 0
    failed = 0
    
    print("=" * 60)
    print("RUNNING AUTOMATED TEST SUITE")
    print("=" * 60)

    for mod in modules:
        for attr in dir(mod):
            if attr.startswith("test_") and callable(getattr(mod, attr)):
                fn = getattr(mod, attr)
                try:
                    fn()
                    print(f"  [PASS] {mod.__name__}.{attr}")
                    passed += 1
                except Exception as exc:
                    print(f"  [FAIL] {mod.__name__}.{attr}: {exc}")
                    traceback.print_exc()
                    failed += 1

    # Run test_sales (unittest suite)
    suite = unittest.defaultTestLoader.loadTestsFromModule(test_sales)
    runner = unittest.TextTestRunner(verbosity=1)
    res = runner.run(suite)
    passed += (res.testsRun - len(res.failures) - len(res.errors))
    failed += (len(res.failures) + len(res.errors))

    print("=" * 60)
    print(f"TEST RESULTS: {passed} passed, {failed} failed")
    print("=" * 60)
    return 0 if failed == 0 else 1

if __name__ == "__main__":
    sys.exit(run_tests())
