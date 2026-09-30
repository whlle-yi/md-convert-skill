"""pytest 共享配置：把 scripts/ 加入 sys.path，使测试可直接导入 md_convert。"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
