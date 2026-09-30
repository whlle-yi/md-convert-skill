"""pytest 共享配置：把技能脚本目录加入 sys.path，使测试可直接导入入口脚本与共享层。"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "skills" / "md-convert" / "scripts"))

# temp/ 不入库（gitignore），而 pytest 9 的 --basetemp 不自动创建父目录，
# CI 全新检出时会因目录缺失而挂——在用到这里之前先建好。
(ROOT / "temp" / "pytest" / "tmp").mkdir(parents=True, exist_ok=True)
