"""PPT 생성기 전역 설정 모듈.

파일 경로, 템플릿 태그, 이미지 검출 파라미터를 한 곳에서 관리한다.
여기 값을 바꾸면 detectors/processors/ppt 모듈 전체에 동일하게 적용된다.
"""

from pathlib import Path

# =====================================
# 기본 경로
# =====================================

# src 폴더 안에 있으므로 한 단계 위가 프로젝트 루트다.
BASE_DIR = Path(__file__).parent.parent

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

# 완성된 PPT는 프로젝트 루트의 output.pptx로 저장한다.
OUTPUT_DIR = (
    BASE_DIR
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

# 템플릿에서 결과 슬라이드와 이미지 자리표시자를 찾는 데 쓰는 문자열 태그.
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

# 범례는 보통 이미지 왼쪽에 세로로 있으므로 왼쪽 60% 영역만 탐색한다.
LEGEND_SEARCH_RATIO = 0.6

# 범례 크롭 여백: 글자가 잘리지 않게 왼쪽을 넉넉히,
# 오른쪽은 수치 라벨까지 포함하되 최대 180px로 제한한다.
LEGEND_MARGIN_LEFT = 0.3

LEGEND_MARGIN_RIGHT = 5.0

LEGEND_MARGIN_TOP = 0.05

LEGEND_MARGIN_BOTTOM = 0.05

LEGEND_MAX_RIGHT_MARGIN = 180

# =====================================
# 모델 크롭
# =====================================

# 모델 크롭 가장자리에 남길 최소 여백(px).
MODEL_MARGIN_LEFT = 1

MODEL_MARGIN_RIGHT = 1

MODEL_MARGIN_TOP = 1

MODEL_MARGIN_BOTTOM = 1

# =====================================
# Debug
# =====================================

# True로 바꾸면 검출 중간 과정을 확인할 수 있다.
DEBUG = False