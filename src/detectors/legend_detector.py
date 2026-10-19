"""해석 결과 이미지에서 색상 범례(legend) 영역을 찾아 잘라내는 모듈."""

import cv2
import numpy as np
import os

from .. import config


class LegendDetector:

    """이미지 왼쪽 영역에서 세로로 긴 범례 블록을 점수화해 찾아낸다."""

    def detect(
        self,
        img,
        base_name,
        temp_dir
    ):

        """가장 점수가 높은 범례 후보를 크롭해 PNG로 저장하고 경로를 반환한다."""

        img_h, img_w = img.shape[:2]

        # =================================
        # 범례 탐색 영역
        # =================================

        # 범례는 대개 이미지 왼쪽에 세로로 배치되므로 왼쪽 일부만 탐색한다.
        roi_width = min(
            img_w,
            max(
                300,
                int(
                    img_w *
                    config.LEGEND_SEARCH_RATIO
                )
            )
        )

        roi = img[:, :roi_width]

        gray = cv2.cvtColor(
            roi,
            cv2.COLOR_BGR2GRAY
        )

        _, binary = cv2.threshold(
            gray,
            240,
            255,
            cv2.THRESH_BINARY_INV
        )

        kernel = np.ones(
            (3, 3),
            np.uint8
        )

        # 작은 점 노이즈를 제거해 범례 블록만 남긴다.
        binary = cv2.morphologyEx(
            binary,
            cv2.MORPH_OPEN,
            kernel
        )

        num_labels, _, stats, _ = (
            cv2.connectedComponentsWithStats(
                binary,
                connectivity=8
            )
        )

        best_score = float("-inf")
        best_rect = None

        # =================================
        # 후보 평가
        # =================================

        for i in range(
            1,
            num_labels
        ):

            x = stats[
                i,
                cv2.CC_STAT_LEFT
            ]

            y = stats[
                i,
                cv2.CC_STAT_TOP
            ]

            w = stats[
                i,
                cv2.CC_STAT_WIDTH
            ]

            h = stats[
                i,
                cv2.CC_STAT_HEIGHT
            ]

            area = stats[
                i,
                cv2.CC_STAT_AREA
            ]

            if area < 30:
                continue

            ratio = h / max(w, 1)

            # 세로로 길수록, 높이가 클수록, 왼쪽에 가까울수록 높은 점수를 준다.
            score = (
                ratio * 1000
                + h * 5
                + max(
                    0,
                    120 - w
                ) * 10
                - x * 2
                - area * 0.05
            )

            # 너무 넓은 객체 제외

            if w > 200:

                score -= 5000

            # 가로형 객체 제외

            if ratio < 2:

                score -= 3000

            if score > best_score:

                best_score = score

                best_rect = (
                    x,
                    y,
                    w,
                    h
                )

        # =================================
        # 실패
        # =================================

        if best_rect is None:

            print(
                f"범례 검출 실패 : "
                f"{base_name}"
            )

            return None

        x, y, w, h = best_rect

        print(
            f"[범례] "
            f"{base_name}"
        )

        # =================================
        # 크롭 영역 계산
        # =================================

        # 범례 주변 여백을 포함해 글자와 색상 막대가 잘리지 않게 한다.
        margin_left = int(
            w *
            config.LEGEND_MARGIN_LEFT
        )

        margin_right = min(
            int(
                w *
                config.LEGEND_MARGIN_RIGHT
            ),
            config.LEGEND_MAX_RIGHT_MARGIN
        )

        margin_top = int(
            h *
            config.LEGEND_MARGIN_TOP
        )

        margin_bottom = int(
            h *
            config.LEGEND_MARGIN_BOTTOM
        )

        x1 = max(
            0,
            x - margin_left
        )

        y1 = max(
            0,
            y - margin_top
        )

        x2 = min(
            img.shape[1],
            x + w + margin_right
        )

        y2 = min(
            img.shape[0],
            y + h + margin_bottom
        )

        legend = img[
            y1:y2,
            x1:x2
        ]

        # =================================
        # 저장
        # =================================

        legend_file = os.path.join(
            temp_dir,
            f"{base_name}_legend.png"
        )

        cv2.imwrite(
            legend_file,
            legend
        )

        return legend_file