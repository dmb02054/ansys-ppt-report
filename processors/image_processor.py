import cv2
import os

from legend_detector import (
    LegendDetector
)


class ImageProcessor:

    def __init__(self):

        self.legend_detector = (
            LegendDetector()
        )

    def process(self, file_path):

        img = cv2.imread(file_path)

        if img is None:
            return None

        base_name = os.path.splitext(
            os.path.basename(file_path)
        )[0]

        legend_file = (
            self.legend_detector.detect(
                img,
                base_name
            )
        )

        model_file = (
            self.extract_model(
                img,
                base_name
            )
        )

        return {
            "title": base_name,
            "model_file": model_file,
            "legend_file": legend_file
        }

    def extract_model(
        self,
        img,
        base_name
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

        contours = sorted(
            contours,
            key=cv2.contourArea,
            reverse=True
        )

        x, y, w, h = cv2.boundingRect(
            contours[0]
        )

        model = img[
            y:y+h,
            x:x+w
        ]

        model_file = (
            f"{base_name}_model.png"
        )

        cv2.imwrite(
            model_file,
            model
        )

        return model_file