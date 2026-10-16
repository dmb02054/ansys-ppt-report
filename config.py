from pathlib import Path
from pptx.util import Inches

# =====================================
# 기본 경로
# =====================================

BASE_DIR = Path(__file__).parent

# =====================================
# 템플릿
# =====================================

TEMPLATE_DIR = (
    BASE_DIR / "templates"
)

TEMPLATE_FILE = (
    TEMPLATE_DIR /
    "report_template.pptx"
)

# =====================================
# 출력
# =====================================

OUTPUT_DIR = (
    BASE_DIR / "output"
)

OUTPUT_DIR.mkdir(
    exist_ok=True
)

OUTPUT_PPT = (
    OUTPUT_DIR /
    "output.pptx"
)

# =====================================
# 결과 슬라이드 태그
# =====================================

RESULT_TEMPLATE_TAG = (
    "[TEMPLATE_RESULT]"
)

TITLE_TAG = (
    "[CASE_TITLE]"
)

# =====================================
# Rack
# =====================================

RACK_MODEL_TAG = (
    "[RACK_MODEL]"
)

RACK_LEGEND_TAG = (
    "[RACK_LEGEND]"
)

# =====================================
# Glass
# =====================================

GLASS_MODEL_TAG = (
    "[GLASS_MODEL]"
)

GLASS_LEGEND_TAG = (
    "[GLASS_LEGEND]"
)

# =====================================
# 범례 검출
# =====================================

LEGEND_SEARCH_RATIO = 0.35

LEGEND_MARGIN_LEFT = 0.3

LEGEND_MARGIN_RIGHT = 5.0

LEGEND_MARGIN_TOP = 0.05

LEGEND_MARGIN_BOTTOM = 0.05

LEGEND_MAX_RIGHT_MARGIN = 180

# =====================================
# 모델 크롭
# =====================================

MODEL_MARGIN_LEFT = 1

MODEL_MARGIN_RIGHT = 1

MODEL_MARGIN_TOP = 1

MODEL_MARGIN_BOTTOM = 1

# =====================================
# Debug
# =====================================

DEBUG = False