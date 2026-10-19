"""프로그램 진입점 모듈.

실제 코드는 src 패키지에 있고, 이 파일은 GUI 실행과
워크플로 함수 다시보내기(re-export)만 담당한다.
"""

from src.workflow import (
    build_cases,
    build_cases_from_assignments,
    build_legend_values,
    generate_presentation
)


def main():
    """Tkinter 루트 창을 만들어 보고서 생성 GUI를 실행한다."""
    import tkinter as tk

    from src.gui import ReportGeneratorApp

    ReportGeneratorApp(tk.Tk()).run()


if __name__ == "__main__":
    main()
