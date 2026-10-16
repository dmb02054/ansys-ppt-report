from copy import deepcopy

from pptx import Presentation

import config


class PPTBuilder:

    def __init__(self):

        self.prs = Presentation(
            str(
                config.TEMPLATE_FILE
            )
        )

        self.result_template_index = (
            self.find_result_template()
        )

    # =====================================
    # 결과 템플릿 찾기
    # =====================================

    def find_result_template(self):

        for idx, slide in enumerate(
            self.prs.slides
        ):

            for shape in slide.shapes:

                if not hasattr(
                    shape,
                    "text"
                ):
                    continue

                if (
                    config.RESULT_TEMPLATE_TAG
                    in shape.text
                ):

                    return idx

        raise ValueError(
            f"{config.RESULT_TEMPLATE_TAG}"
            f" 슬라이드를 찾을 수 없습니다."
        )

    # =====================================
    # 슬라이드 복제
    # =====================================

    def duplicate_slide(
        self,
        slide_index
    ):

        source = (
            self.prs.slides[
                slide_index
            ]
        )

        new_slide = (
            self.prs.slides.add_slide(
                source.slide_layout
            )
        )

        for shape in source.shapes:

            new_slide.shapes._spTree.insert_element_before(
                deepcopy(
                    shape.element
                ),
                "p:extLst"
            )

        return new_slide

    # =====================================
    # 이미지 자리표시자 교체
    # =====================================

    def replace_placeholder_image(
        self,
        slide,
        placeholder,
        image_file
    ):

        if not image_file:
            return

        for shape in slide.shapes:

            if not hasattr(
                shape,
                "text"
            ):
                continue

            if (
                shape.text.strip()
                != placeholder
            ):
                continue

            left = shape.left
            top = shape.top

            height = shape.height

            slide.shapes._spTree.remove(
                shape.element
            )

            # 높이만 지정
            slide.shapes.add_picture(
                image_file,
                left,
                top,
                height=height
            )

            break

    # =====================================
    # 결과 슬라이드 생성
    # =====================================

    def add_case_slide(
        self,
        case_name,
        rack,
        glass
    ):

        slide = self.duplicate_slide(
            self.result_template_index
        )

        for shape in slide.shapes:

            if not hasattr(
                shape,
                "text"
            ):
                continue

            text = (
                shape.text.strip()
            )

            # -------------------
            # Case 제목
            # -------------------

            if (
                text
                == config.TITLE_TAG
            ):

                shape.text = (
                    case_name
                )

            # -------------------
            # 결과 템플릿 태그 삭제
            # -------------------

            elif (
                text
                == config.RESULT_TEMPLATE_TAG
            ):

                shape.text = ""

        # -----------------------
        # Rack
        # -----------------------

        if rack:

            self.replace_placeholder_image(
                slide,
                "[RACK_MODEL]",
                rack["model_file"]
            )

            self.replace_placeholder_image(
                slide,
                "[RACK_LEGEND]",
                rack["legend_file"]
            )

        # -----------------------
        # Glass
        # -----------------------

        if glass:

            self.replace_placeholder_image(
                slide,
                "[GLASS_MODEL]",
                glass["model_file"]
            )

            self.replace_placeholder_image(
                slide,
                "[GLASS_LEGEND]",
                glass["legend_file"]
            )

    # =====================================
    # 슬라이드 삭제
    # =====================================

    def remove_slide(
        self,
        slide_index
    ):

        slide_id = (
            self.prs.slides._sldIdLst[
                slide_index
            ]
        )

        self.prs.part.drop_rel(
            slide_id.rId
        )

        del self.prs.slides._sldIdLst[
            slide_index
        ]

    # =====================================
    # 저장
    # =====================================

    def save(self):

        self.remove_slide(
            self.result_template_index
        )

        self.prs.save(
            str(
                config.OUTPUT_PPT
            )
        )