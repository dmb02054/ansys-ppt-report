import cv2
import os


class ModelDetector:

    def detect(
        self,
        img,
        base_name,
        temp_dir
    ):

        gray = cv2.cvtColor(
            img,
            cv2.COLOR_BGR2GRAY
        )

        _, mask = cv2.threshold(
            gray,
            250,
            255,
            cv2.THRESH_BINARY_INV
        )

        contours, _ = cv2.findContours(
            mask,
            cv2.RETR_EXTERNAL,
            cv2.CHAIN_APPROX_SIMPLE
        )

        if not contours:

            print(
                f"모델 검출 실패 : {base_name}"
            )

            return None

        img_area = (
            img.shape[0]
            * img.shape[1]
        )

        contours = sorted(
            contours,
            key=cv2.contourArea,
            reverse=True
        )

        model_cnt = None

        for cnt in contours:

            if (
                cv2.contourArea(cnt)
                > img_area * 0.05
            ):

                model_cnt = cnt

                break

        if model_cnt is None:

            model_cnt = contours[0]

        x, y, w, h = cv2.boundingRect(
            model_cnt
        )

        # -------------------
        # 모델 여백
        # -------------------

        margin_left = 1

        margin_right = 1

        margin_top = 1

        margin_bottom = 1

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