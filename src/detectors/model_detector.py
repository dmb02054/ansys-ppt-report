"""해석 결과 이미지에서 해석 모델(본체) 영역을 찾아 잘라내는 모듈."""

import cv2
import os

import numpy as np


class ModelDetector:

    """여러 밝기 임계값으로 만든 마스크를 합쳐 모델 윤곽을 찾아낸다."""

    def detect(
        self,
        img,
        base_name,
        temp_dir
    ):

        """모델로 가장 적합한 윤곽을 크롭해 PNG로 저장하고 경로를 반환한다."""

        if img is None or img.size == 0:
            print(
                f"모델 검출 실패 : {base_name}"
            )
            return None

        img_h, img_w = img.shape[:2]
        img_area = img_h * img_w

        masks = []

        # 밝기 기준을 여러 개 써서 조명·배경 차이로 모델이 누락되지 않게 한다.
        for threshold in (200, 220, 240, 250):
            gray = cv2.cvtColor(
                img,
                cv2.COLOR_BGR2GRAY
            )

            _, mask = cv2.threshold(
                gray,
                threshold,
                255,
                cv2.THRESH_BINARY_INV
            )

            masks.append(mask)

        final_mask = np.zeros_like(
            masks[0],
            dtype=np.uint8
        )

        for mask in masks:
            final_mask = cv2.bitwise_or(
                final_mask,
                mask
            )

        kernel = np.ones(
            (3, 3),
            np.uint8
        )

        # 끊어진 모델 외곽선을 연결해 하나의 덩어리로 만든다.
        final_mask = cv2.morphologyEx(
            final_mask,
            cv2.MORPH_CLOSE,
            kernel
        )

        contours, _ = cv2.findContours(
            final_mask,
            cv2.RETR_EXTERNAL,
            cv2.CHAIN_APPROX_SIMPLE
        )

        if not contours:
            print(
                f"모델 검출 실패 : {base_name}"
            )
            return None

        best_cnt = None
        best_score = float("-inf")

        for cnt in contours:
            area = cv2.contourArea(cnt)
            if area < img_area * 0.02:
                continue

            x, y, w, h = cv2.boundingRect(cnt)
            if w <= 0 or h <= 0:
                continue

            if w > img_w * 0.9:
                continue

            if h < img_h * 0.12:
                continue

            cx = x + w / 2
            cy = y + h / 2
            center_distance = abs(cx - img_w / 2)
            aspect_ratio = h / max(w, 1)
            # 면적이 크고, 세로로 길고, 화면 중앙에 가까운 윤곽을 모델로 선택한다.
            score = (
                area
                + h * 30
                + aspect_ratio * 5000
                - center_distance * 0.5
            )

            # 왼쪽에 붙은 좁은 객체는 범례일 가능성이 높아 감점한다.
            if x < img_w * 0.15 and w < img_w * 0.45:
                score -= area * 0.3

            if score > best_score:
                best_score = score
                best_cnt = cnt

        # 점수로 선택하지 못하면 가장 큰 윤곽을 모델로 사용한다.
        if best_cnt is None:
            best_cnt = max(
                contours,
                key=cv2.contourArea
            )

        x, y, w, h = cv2.boundingRect(
            best_cnt
        )

        margin_left = 2
        margin_right = 2
        margin_top = 2
        margin_bottom = 2

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

        model = img[
            y1:y2,
            x1:x2
        ]

        model_file = os.path.join(
            temp_dir,
            f"{base_name}_model.png"
        )

        cv2.imwrite(
            model_file,
            model
        )

        print(
            f"[모델] {base_name}"
        )

        return model_file