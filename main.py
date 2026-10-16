from tkinter import (
    Tk,
    filedialog
)

from image_processor import (
    ImageProcessor
)

from ppt_builder import (
    PPTBuilder
)


def select_files():

    root = Tk()

    root.withdraw()

    root.attributes(
        "-topmost",
        True
    )

    return (
        filedialog.askopenfilenames(
            title="이미지 선택",
            filetypes=[
                (
                    "Image Files",
                    "*.png *.jpg *.jpeg *.bmp"
                )
            ]
        )
    )


def main():

    files = select_files()

    if not files:

        print(
            "선택된 파일이 없습니다."
        )

        return

    processor = (
        ImageProcessor()
    )

    ppt = (
        PPTBuilder()
    )

    for file_path in files:

        result = (
            processor.process(
                file_path
            )
        )

        if result:

            ppt.add_slide(
                result
            )

    ppt.save()

    print(
        "PPT 생성 완료"
    )


if __name__ == "__main__":

    main()