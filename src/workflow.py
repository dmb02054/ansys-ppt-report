"""보고서 읽기, 이미지 처리, PPT 조립을 순서대로 지휘하는 모듈.

GUI나 명령행에서 받은 입력을 검증·정규화한 뒤 ImageProcessor로
이미지를 자르고 PPTBuilder로 슬라이드를 채운다.
"""

import re
import tempfile
from pathlib import Path

from . import config
from .processors.image_processor import ImageProcessor
from .processors.report_reader import read_report_data
from .ppt.ppt_builder import PPTBuilder


def build_cases(files):

    """파일명에서 Case·부품을 읽어 {Case: {부품: 경로}} 형태로 묶는다."""

    cases = {}

    for file_path in files:
        metadata = parse_image_metadata(file_path)
        case_name = metadata["case"]
        part_name = metadata["part"]
        if not case_name or not part_name:
            continue

        cases.setdefault(case_name, {})[part_name] = file_path

    return cases


def parse_image_metadata(file_path):

    """파일명에서 Case 이름, 부품(rack/glass), Max/Min 값을 추출한다."""

    # 생성된 파일명의 Case/부품/Max/Min 토큰을 읽는다.
    # 읽지 못한 값은 빈 문자열로 두고 GUI에서 직접 수정할 수 있다.

    stem = Path(file_path).stem
    role_match = re.search(
        r"(?:^|[_\s-])(rack|glass)(?=$|[_\s-])",
        stem,
        flags=re.IGNORECASE
    )
    case_name = stem[:role_match.start()].strip(" _-") if role_match else ""
    role = role_match.group(1).lower() if role_match else ""
    number = r"([+-]?(?:\d+(?:\.\d*)?|\.\d+))"
    max_match = re.search(
        rf"(?:max(?:imum)?|최대)\s*(?:[:=_]\s*)?{number}",
        stem,
        flags=re.IGNORECASE
    )
    min_match = re.search(
        rf"(?:min(?:imum)?|최소)\s*(?:[:=_]\s*)?{number}",
        stem,
        flags=re.IGNORECASE
    )

    return {
        "case": case_name,
        "part": role,
        "max": max_match.group(1) if max_match else "",
        "min": min_match.group(1) if min_match else ""
    }


def build_cases_from_assignments(assignments):

    """GUI에서 사용자가 직접 배정한 Case/부품 맵을 정규화한다."""

    # 파일명 추측보다 사용자가 GUI에서 명시한 배정을 항상 우선한다.

    cases = {}

    for case_name, parts in assignments.items():
        normalized_name = str(case_name).strip()
        if not normalized_name:
            continue

        normalized_parts = {
            str(part).lower(): file_path
            for part, file_path in parts.items()
            if str(part).lower() in ("rack", "glass") and file_path
        }
        if normalized_parts:
            cases[normalized_name] = normalized_parts

    return cases


def build_legend_values(image_roles, values_by_image):

    """이미지별로 입력된 Max/Min 값을 Case/부품별 구조로 모은다."""

    # 이미지 경로 기준의 GUI 입력값을 슬라이드에서 쓰는 Case/부품 구조로 바꾼다.

    values_by_case = {}
    for image_file, (case_name, part) in image_roles.items():
        values = values_by_image.get(image_file, {})
        values_by_case.setdefault(case_name, {})[part] = {
            "max": str(values.get("max", "")).strip(),
            "min": str(values.get("min", "")).strip()
        }

    return values_by_case


def rename_assigned_images(assignments, legend_values):

    """PPT 생성이 끝난 뒤 원본 이미지 파일명을 'Case_부품_Max/Min' 형식으로 바꾼다.

    같은 이름의 파일이 이미 있으면 '_2', '_3'처럼 번호를 붙여
    기존 파일을 덮어쓰지 않는다.
    반환: ({이전 경로: 새 경로}, [오류 메시지])
    """

    renamed = {}
    errors = []

    for case_name, parts in assignments.items():
        for part, file_path in parts.items():
            source = Path(file_path)
            if not source.is_file():
                errors.append(f"원본 이미지가 없습니다: {source}")
                continue

            values = legend_values.get(case_name, {}).get(part, {})
            # 입력된 Max/Min이 있으면 파일명에 함께 포함한다.
            filename_parts = [case_name, part.title()]
            if values.get("max"):
                filename_parts.extend(("Max", str(values["max"]).strip()))
            if values.get("min"):
                filename_parts.extend(("Min", str(values["min"]).strip()))

            stem = "_".join(filename_parts)
            # Windows 파일명에 쓸 수 없는 문자는 밑줄로 바꾼다.
            stem = re.sub(r'[<>:"/\\|?*]+', "_", stem)
            stem = stem.strip(" ._")
            target = source.with_name(stem + source.suffix)
            if target.resolve() == source.resolve():
                continue

            # 같은 이름의 파일이 이미 있으면 번호를 붙여 덮어쓰기를 막는다.
            collision_index = 2
            while target.exists():
                target = source.with_name(
                    f"{stem}_{collision_index}{source.suffix}"
                )
                collision_index += 1

            try:
                source.rename(target)
                renamed[str(source)] = str(target)
            except OSError as error:
                errors.append(f"파일명 변경 실패: {source} ({error})")

    return renamed, errors


def generate_presentation(
    report_file,
    files,
    progress_callback=None,
    case_order=None,
    report_overrides=None,
    case_assignments=None,
    legend_values=None,
    output_path=None
):

    """보고서(선택)와 배정된 결과 이미지로 PPT를 만들어 저장한다.

    output_path가 없으면 config의 기본 출력 경로에 저장한다.
    반환: (출력 파일 경로, 생성된 Case 수)
    """

    if not files:
        raise ValueError("해석 결과 이미지를 선택해 주세요.")

    # GUI는 Case/부품 배정 맵을 넘긴다. 파일명 해석은 명령행 사용 때만 쓴다.
    cases = (
        build_cases_from_assignments(case_assignments)
        if case_assignments is not None
        else build_cases(files)
    )
    if not cases:
        raise ValueError(
            "이미지를 선택한 Case에 Glass 또는 Rack으로 지정해 주세요."
        )

    # GUI에서 정한 Case 순서를 우선 적용하고, 빠진 Case는 뒤에 붙인다.
    if case_order:
        ordered_names = [
            case_name
            for case_name in case_order
            if case_name in cases
        ]
        ordered_names.extend(
            case_name
            for case_name in cases
            if case_name not in ordered_names
        )
        cases = {
            case_name: cases[case_name]
            for case_name in ordered_names
        }

    notify = progress_callback or (lambda message: None)
    # 보고서 파일은 선택 사항이다. 없으면 표지 안내 문구만 지운다.
    if report_file:
        notify("완료보고서 정보를 읽는 중입니다.")
        report_data = read_report_data(report_file)
    else:
        notify("표지 정보 없이 결과 슬라이드를 준비합니다.")
        report_data = {}

    report_fields = (
        "title",
        "analysis_number",
        "contact",
        "start_date",
        "end_date",
        "agenda",
        "background",
        "objective",
        "conclusion"
    )
    for field in report_fields:
        value = (report_overrides or {}).get(field)
        if value and str(value).strip():
            report_data[field] = str(value).strip()
    report_data["skip_cover"] = not any(
        report_data.get(field)
        for field in (
            "title",
            "analysis_number",
            "contact",
            "start_date",
            "end_date"
        )
    )

    # 크롭 이미지는 임시 폴더에만 두고, 최종 PPT에 삽입된 뒤 함께 삭제된다.
    with tempfile.TemporaryDirectory() as temp_dir:
        processor = ImageProcessor(temp_dir)
        presentation = PPTBuilder()
        presentation.fill_report_info(report_data)

        for case_name, case_data in cases.items():
            notify(f"{case_name} 해석 이미지를 처리하는 중입니다.")
            rack = None
            glass = None

            if "rack" in case_data:
                rack = processor.process(case_data["rack"])
            if "glass" in case_data:
                glass = processor.process(case_data["glass"])

            if rack is None and glass is None:
                raise RuntimeError(
                    f"{case_name} 이미지에서 모델을 처리하지 못했습니다."
                )

            presentation.add_case_slide(
                case_name,
                rack,
                glass,
                legend_values=(legend_values or {}).get(case_name, {})
            )

        # 편집용 템플릿 슬라이드는 저장 때 제거되고, 마지막에 정리 표 슬라이드를 붙인다.
        presentation.add_case_summary_slide(
            list(cases),
            legend_values or {}
        )
        notify("PowerPoint 파일을 저장하는 중입니다.")
        presentation.save(output_path)

    # 저장 위치를 지정하지 않으면 config의 기본 출력 경로를 돌려준다.
    if output_path:
        return Path(output_path), len(cases)
    return config.OUTPUT_PPT, len(cases)
