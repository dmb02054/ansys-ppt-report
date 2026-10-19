"""원본 해석 이미지를 범례와 모델 이미지로 분리하는 모듈."""

import cv2
import os

from ..detectors.legend_detector import (
    LegendDetector
)

from ..detectors.model_detector import (
    ModelDetector
)


class ImageProcessor:

    """LegendDetector와 ModelDetector를 묶어 이미지 한 장을 처리한다."""

    def __init__(
        self,
        temp_dir
    ):

        """크롭 결과를 저장할 임시 폴더 경로를 받아 검출기를 준비한다."""

        self.temp_dir = temp_dir

        self.legend_detector = (
            LegendDetector()
        )

        self.model_detector = (
            ModelDetector()
        )

    def process(
        self,
        file_path
    ):

        """이미지를 읽어 범례·모델 파일 경로를 dict로 반환한다. 실패 시 None."""

        img = cv2.imread(
            file_path
        )

        if img is None:

            print(
                f"이미지 로드 실패 : {file_path}"
            )

            return None

        base_name = os.path.splitext(
            os.path.basename(
                file_path
            )
        )[0]

        # =================================
        # 범례 추출
        # =================================

        legend_file = (
            self.legend_detector.detect(
                img,
                base_name,
                self.temp_dir
            )
        )

        # =================================
        # 모델 추출
        # =================================

        model_file = (
            self.model_detector.detect(
                img,
                base_name,
                self.temp_dir
            )
        )

        if model_file is None:

            print(
                f"모델 검출 실패 : "
                f"{base_name}"
            )

            return None

        # =================================
        # 결과 반환
        # =================================

        return {
            "title": base_name,

            "model_file": model_file,

            "legend_file": legend_file
        }