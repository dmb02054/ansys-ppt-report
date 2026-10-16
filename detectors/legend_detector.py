import cv2
import numpy as np


class LegendDetector:

    def detect(self, img, base_name):

        img_h, img_w = img.shape[:2]

        roi_width = int(img_w * 0.35)

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

        kernel = np.ones((3, 3), np.uint8)

        binary = cv2.morphologyEx(
            binary,
            cv2.MORPH_OPEN,
            kernel
        )

        num_labels, labels, stats, _ = (
            cv2.connectedComponentsWithStats(
                binary,
                connectivity=8
            )
        )

        best_score = -999999
        best_rect = None

        for i in range(1, num_labels):

            x = stats[i, cv2.CC_STAT_LEFT]
            y = stats[i, cv2.CC_STAT_TOP]

            w = stats[i, cv2.CC_STAT_WIDTH]
            h = stats[i, cv2.CC_STAT_HEIGHT]

            area = stats[i, cv2.CC_STAT_AREA]

            if area < 30:
                continue

            ratio = h / max(w, 1)

            score = (
                ratio * 1000
                + h * 5
                + max(0, 120 - w) * 10
                - x * 2
                - area * 0.05
            )

            if score > best_score:

                best_score = score
                best_rect = (
                    x,
                    y,
                    w,
                    h
                )

        if best_rect is None:
            return None

        x, y, w, h = best_rect

        margin_left = int(w * 0.3)
        margin_right = int(w * 5)

        margin_top = int(h * 0.05)
        margin_bottom = int(h * 0.05)

        x1 = max(0, x - margin_left)
        y1 = max(0, y - margin_top)

        x2 = min(
            img_w,
            x + w + margin_right
        )

        y2 = min(
            img_h,
            y + h + margin_bottom
        )

        legend = img[
            y1:y2,
            x1:x2
        ]

        legend_file = (
            f"{base_name}_legend.png"
        )

        cv2.imwrite(
            legend_file,
            legend
        )

        return legend_file