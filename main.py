import tempfile
import shutil
import os

from tkinter import (
    Tk,
    filedialog
)

from processors.image_processor import (
    ImageProcessor
)

from ppt.ppt_builder import (
    PPTBuilder
)


def select_files():

    root = Tk()

    root.withdraw()

    root.attributes(
        "-topmost",
        True
    )

    return filedialog.askopenfilenames(
        title="해석 결과 이미지 선택",
        filetypes=[
            (
                "Image Files",
                "*.png *.jpg *.jpeg *.bmp"
            )
        ]
    )


def build_cases(files):

    cases = {}

    for file_path in files:

        filename = os.path.basename(
            file_path
        )

        name = os.path.splitext(
            filename
        )[0]

        if "_Rack" in name:

            case_name = (
                name.replace(
                    "_Rack",
                    ""
                )
            )

            cases.setdefault(
                case_name,
                {}
            )

            cases[case_name][
                "rack"
            ] = file_path

        elif "_Glass" in name:

            case_name = (
                name.replace(
                    "_Glass",
                    ""
                )
            )

            cases.setdefault(
                case_name,
                {}
            )

            cases[case_name][
                "glass"
            ] = file_path

    return cases


def main():

    files = select_files()

    if not files:

        return

    temp_dir = tempfile.mkdtemp()

    try:

        processor = ImageProcessor(
            temp_dir
        )

        ppt = PPTBuilder()

        cases = build_cases(
            files
        )

        for case_name, case_data in cases.items():

            rack = None
            glass = None

            if "rack" in case_data:

                rack = processor.process(
                    case_data["rack"]
                )

            if "glass" in case_data:

                glass = processor.process(
                    case_data["glass"]
                )

            ppt.add_case_slide(
                case_name,
                rack,
                glass
            )

        ppt.save()

    finally:

        shutil.rmtree(
            temp_dir,
            ignore_errors=True
        )


if __name__ == "__main__":
    main()
