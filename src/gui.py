"""해석 결과 PPT 생성기의 Tkinter GUI 모듈.

보고서 선택, 이미지 목록 관리, 미리보기, Case 배정,
범례 값 입력, PPT 생성 실행을 담당한다.
"""

import os
import queue
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from . import config
from .processors.image_processor import ImageProcessor
from .processors.report_reader import read_report_data
from .workflow import (
    build_cases_from_assignments,
    build_legend_values,
    generate_presentation,
    parse_image_metadata,
    rename_assigned_images
)


# 표지 편집창에 표시할 항목: (화면 라벨, 보고서 데이터 키)
REPORT_TEXT_FIELDS = (
    ("제목", "title"),
    ("해석 번호", "analysis_number"),
    ("작성자", "contact"),
    ("수행 시작", "start_date"),
    ("수행 종료", "end_date"),
    ("안건명", "agenda"),
    ("배경", "background"),
    ("목표", "objective"),
    ("최종 결론", "conclusion")
)


def format_report_text(values):

    """보고서 데이터를 '라벨 : 값' 형태의 편집용 텍스트로 만든다."""

    return "\n".join(
        f"{label} : {values.get(key, '') or ''}"
        for label, key in REPORT_TEXT_FIELDS
    )


def parse_report_text(text):

    """편집창 텍스트를 다시 항목별 dict로 읽는다.

    '라벨 :'로 시작하지 않는 줄은 직전 항목의 이어지는 내용으로 합친다.
    """

    fields_by_label = dict(REPORT_TEXT_FIELDS)
    values = {key: "" for _, key in REPORT_TEXT_FIELDS}
    current_key = None

    for line in text.splitlines():
        label, separator, value = line.partition(":")
        key = fields_by_label.get(label.strip()) if separator else None
        if key:
            current_key = key
            values[key] = value.strip()
        elif current_key and line.strip():
            values[current_key] += "\n" + line.strip()

    start_date = values["start_date"]
    end_date = values["end_date"]
    values["period"] = f"{start_date} ~ {end_date}".strip(" ~")
    return values


class ReportGeneratorApp:
    """이미지 선택·역할 지정·PPT 생성을 처리하는 Tkinter 메인 창."""

    def __init__(self, root):
        """창 크기·상태 변수를 초기화하고 UI를 구성한다."""
        self.root = root
        self.root.title("Glass 처짐 해석 보고서 작성")
        self.root.geometry("1240x800")
        self.root.minsize(900, 650)

        # ---- 선택 상태: 보고서/이미지/Case 배정 ----
        self.report_file = ""
        self.image_files = ()
        self.visible_image_files = ()
        self.case_assignments = {}
        self.image_roles = {}
        self.case_order = []
        # ---- 미리보기 상태 ----
        self.preview_photo = None
        self.preview_original = None
        self.preview_results = {}
        self.preview_photos = {}
        self.legend_values_by_image = {}
        self.loading_legend_values = False
        self.preview_temp_dir = tempfile.TemporaryDirectory()
        self.current_preview_file = None
        # ---- 생성 작업 상태: 백그라운드 스레드와 이벤트 큐 ----
        self.is_busy = False
        self.events = queue.Queue()

        self.report_summary = tk.StringVar(
            value="선택된 보고서가 없습니다."
        )
        self.image_summary = tk.StringVar(
            value="선택된 이미지가 없습니다."
        )
        self.rename_original_files = tk.BooleanVar(value=True)
        self.image_filter = tk.StringVar(
            value="전체"
        )
        self.case_name = tk.StringVar(
            value="Case 1"
        )
        self.preview_info = tk.StringVar(
            value="목록에서 이미지를 선택하면 미리보기가 표시됩니다."
        )
        self.max_value = tk.StringVar()
        self.min_value = tk.StringVar()
        # 저장 경로·파일명은 사용자가 직접 바꿀 수 있다.
        self.output_path = tk.StringVar(
            value=str(config.OUTPUT_PPT)
        )
        self.status = tk.StringVar(
            value="보고서와 해석 이미지를 선택해 주세요."
        )

        self._build_ui()
        self.max_value.trace_add("write", self._save_legend_values)
        self.min_value.trace_add("write", self._save_legend_values)
        self.root.after(100, self._poll_events)
        self.root.protocol("WM_DELETE_WINDOW", self._close)

    def _build_ui(self):
        """보고서·이미지·Case·생성 버튼 영역을 배치한다."""
        container = ttk.Frame(
            self.root,
            padding=(18, 16)
        )
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        ttk.Label(
            container,
            text="해석 결과 보고서 제작",
            font=("맑은 고딕", 17, "bold")
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            container,
            text="보고서 정보를 확인하고 결과 이미지를 정리한 뒤 PPT를 생성합니다.",
            foreground="#52616b"
        ).grid(row=1, column=0, sticky="w", pady=(3, 12))

        report_frame = ttk.LabelFrame(
            container,
            text="표지 제작",
            padding=(12, 10)
        )
        report_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        report_frame.columnconfigure(2, weight=1)
        report_frame.columnconfigure(4, weight=1)
        self.report_button = ttk.Button(
            report_frame,
            text="파일 선택",
            command=self._choose_report
        )
        self.report_button.grid(
            row=0,
            column=0,
            rowspan=6,
            sticky="nsew",
            padx=(0, 18),
            pady=2
        )
        ttk.Label(
            report_frame,
            textvariable=self.report_summary,
            foreground="#52616b"
        ).grid(row=0, column=1, columnspan=2, sticky="w", pady=(0, 5))
        # 표지 항목 전체를 한 편집창에 넣고 스크롤로 볼 수 있게 한다.
        self.report_text = tk.Text(
            report_frame,
            height=5,
            wrap="word",
            undo=True,
            borderwidth=1,
            relief="solid"
        )
        self.report_text.grid(row=1, column=1, sticky="nsew", pady=(0, 2))
        report_scrollbar = ttk.Scrollbar(
            report_frame,
            orient="vertical",
            command=self.report_text.yview
        )
        report_scrollbar.grid(row=1, column=2, sticky="ns", pady=(0, 2))
        self.report_text.configure(
            yscrollcommand=report_scrollbar.set
        )
        self._set_report_text({})

        image_frame = ttk.LabelFrame(
            container,
            text="Glass 데이터 입력",
            padding=(10, 9)
        )
        image_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        image_frame.columnconfigure(0, weight=3, minsize=300)
        image_frame.columnconfigure(1, weight=4, minsize=250)
        image_frame.columnconfigure(2, weight=3, minsize=300)
        image_frame.rowconfigure(1, weight=1)

        image_toolbar = ttk.Frame(image_frame)
        image_toolbar.grid(row=0, column=0, columnspan=3, sticky="ew")
        image_toolbar.columnconfigure(1, weight=1)
        self.image_button = ttk.Button(
            image_toolbar,
            text="이미지 파일 선택",
            command=self._choose_images
        )
        self.image_button.grid(row=0, column=0, sticky="w")
        ttk.Label(
            image_toolbar,
            textvariable=self.image_summary
        ).grid(row=0, column=1, sticky="e")
        self.label_image_button = ttk.Checkbutton(
            image_toolbar,
            text="PPT 생성 후 원본 파일명 변경",
            variable=self.rename_original_files
        )
        self.label_image_button.grid(
            row=0,
            column=2,
            sticky="e",
            padx=(8, 0)
        )

        image_list_frame = ttk.Frame(image_frame)
        image_list_frame.grid(
            row=1,
            column=0,
            sticky="nsew",
            padx=(0, 8),
            pady=(8, 0)
        )
        image_list_frame.columnconfigure(0, weight=1)
        image_list_frame.rowconfigure(1, weight=1)
        filter_frame = ttk.Frame(image_list_frame)
        filter_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        self.filter_buttons = []
        for column, value in enumerate(("전체", "미지정", "Glass", "Rack")):
            button = ttk.Radiobutton(
                filter_frame,
                text=value,
                value=value,
                variable=self.image_filter,
                command=self._refresh_image_list
            )
            button.grid(row=0, column=column, padx=(0, 5))
            self.filter_buttons.append(button)

        self.image_list = tk.Listbox(
            image_list_frame,
            activestyle="none",
            borderwidth=1,
            relief="solid",
            exportselection=False
        )
        self.image_list.grid(
            row=1,
            column=0,
            sticky="nsew",
            pady=(0, 5)
        )
        image_scrollbar = ttk.Scrollbar(
            image_list_frame,
            orient="vertical",
            command=self.image_list.yview
        )
        image_scrollbar.grid(row=1, column=1, sticky="ns", pady=(0, 5))
        self.image_list.configure(
            yscrollcommand=image_scrollbar.set
        )
        self.image_list.bind(
            "<<ListboxSelect>>",
            self._show_preview
        )

        preview_frame = ttk.LabelFrame(
            image_frame,
            text="선택 이미지 분석 미리보기",
            padding=8
        )
        preview_frame.grid(
            row=1,
            column=1,
            sticky="nsew",
            padx=4,
            pady=(8, 0)
        )
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        preview_panels = ttk.Frame(preview_frame)
        preview_panels.grid(row=0, column=0, sticky="nsew")
        preview_panels.columnconfigure(0, weight=5)
        preview_panels.columnconfigure(1, weight=1)
        preview_panels.rowconfigure(0, weight=1)
        self.preview_canvases = {}
        for column, (key, label) in enumerate((
            ("legend", "범례"),
            ("model", "모델")
        )):
            panel = ttk.LabelFrame(
                preview_panels,
                text=label,
                padding=4
            )
            panel.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=(0, 4) if column == 0 else (4, 0)
            )
            panel.columnconfigure(0, weight=1)
            panel.rowconfigure(0, weight=1)
            canvas = tk.Canvas(
                panel,
                background="white",
                highlightthickness=0
            )
            canvas.grid(row=0, column=0, sticky="nsew")
            canvas.bind(
                "<Configure>",
                lambda _event, preview_key=key: self._render_preview(preview_key)
            )
            self.preview_canvases[key] = canvas

        ttk.Label(
            preview_frame,
            textvariable=self.preview_info,
            anchor="center",
            foreground="#52616b"
        ).grid(row=1, column=0, sticky="ew", pady=(5, 0))

        value_frame = ttk.Frame(preview_frame)
        value_frame.grid(row=2, column=0, sticky="ew", pady=(7, 0))
        value_frame.columnconfigure(1, weight=1)
        value_frame.columnconfigure(3, weight=1)
        ttk.Label(value_frame, text="최대값").grid(
            row=0,
            column=0,
            sticky="e",
            padx=(0, 5)
        )
        self.max_value_entry = ttk.Entry(
            value_frame,
            textvariable=self.max_value,
            width=12
        )
        self.max_value_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12))
        ttk.Label(value_frame, text="최소값").grid(
            row=0,
            column=2,
            sticky="e",
            padx=(0, 5)
        )
        self.min_value_entry = ttk.Entry(
            value_frame,
            textvariable=self.min_value,
            width=12
        )
        self.min_value_entry.grid(row=0, column=3, sticky="ew")

        assignment_controls = ttk.Frame(preview_frame)
        assignment_controls.grid(row=3, column=0, sticky="ew", pady=(7, 0))
        assignment_controls.columnconfigure(0, weight=1)
        assignment_controls.columnconfigure(1, weight=1)
        self.assign_glass_button = ttk.Button(
            assignment_controls,
            text="Glass 지정",
            command=lambda: self._assign_current_image("glass"),
            state="disabled"
        )
        self.assign_glass_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.assign_rack_button = ttk.Button(
            assignment_controls,
            text="Rack 지정",
            command=lambda: self._assign_current_image("rack"),
            state="disabled"
        )
        self.assign_rack_button.grid(row=0, column=1, sticky="ew", padx=(4, 0))

        case_frame = ttk.LabelFrame(
            image_frame,
            text="Case 순서",
            padding=8
        )
        case_frame.grid(
            row=1,
            column=2,
            sticky="nsew",
            padx=(8, 0),
            pady=(8, 0)
        )
        case_frame.columnconfigure(0, weight=1)
        case_frame.rowconfigure(1, weight=1)
        case_entry_frame = ttk.Frame(case_frame)
        case_entry_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        case_entry_frame.columnconfigure(0, weight=1)
        self.case_name_entry = ttk.Entry(
            case_entry_frame,
            textvariable=self.case_name
        )
        self.case_name_entry.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.add_case_button = ttk.Button(
            case_entry_frame,
            text="Case 추가",
            command=self._add_case
        )
        self.add_case_button.grid(row=0, column=1)
        self.case_list = tk.Listbox(
            case_frame,
            activestyle="none",
            borderwidth=1,
            relief="solid",
            exportselection=False
        )
        self.case_list.grid(row=1, column=0, sticky="nsew")
        case_scrollbar = ttk.Scrollbar(
            case_frame,
            orient="vertical",
            command=self.case_list.yview
        )
        case_scrollbar.grid(row=1, column=1, sticky="ns")
        self.case_list.configure(
            yscrollcommand=case_scrollbar.set
        )
        case_controls = ttk.Frame(case_frame)
        case_controls.grid(row=2, column=0, columnspan=2, sticky="e", pady=(6, 0))
        self.move_up_button = ttk.Button(
            case_controls,
            text="▲",
            width=4,
            command=lambda: self._move_case(-1)
        )
        self.move_up_button.pack(side="left", padx=(0, 4))
        self.move_down_button = ttk.Button(
            case_controls,
            text="▼",
            width=4,
            command=lambda: self._move_case(1)
        )
        self.move_down_button.pack(side="left")

        self.input_widgets = [
            self.report_button,
            self.image_button,
            self.label_image_button,
            self.report_text,
            self.case_name_entry,
            self.add_case_button,
            self.image_list,
            self.case_list,
            self.max_value_entry,
            self.min_value_entry,
            self.assign_glass_button,
            self.assign_rack_button,
            self.move_up_button,
            self.move_down_button,
            *self.filter_buttons
        ]

        # 저장 파일명·경로를 직접 입력하거나 대화상자로 고를 수 있게 한다.
        output_frame = ttk.LabelFrame(
            container,
            text="저장 위치",
            padding=(10, 7)
        )
        output_frame.grid(row=4, column=0, sticky="ew", pady=(0, 7))
        output_frame.columnconfigure(0, weight=1)
        self.output_path_entry = ttk.Entry(
            output_frame,
            textvariable=self.output_path
        )
        self.output_path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.output_path_button = ttk.Button(
            output_frame,
            text="저장 위치 변경",
            command=self._choose_output_path
        )
        self.output_path_button.grid(row=0, column=1)
        self.input_widgets.extend((
            self.output_path_entry,
            self.output_path_button
        ))

        activity = ttk.Frame(container)
        activity.grid(row=5, column=0, sticky="ew")
        activity.columnconfigure(0, weight=1)
        ttk.Label(
            activity,
            textvariable=self.status
        ).grid(row=0, column=0, sticky="w")

        actions = ttk.Frame(container)
        actions.grid(row=6, column=0, sticky="e", pady=(9, 0))
        self.generate_button = ttk.Button(
            actions,
            text="파일 생성",
            command=self._start_generation,
            state="disabled"
        )
        self.generate_button.pack(side="right", ipadx=20, ipady=5)

    def _set_report_text(self, values):
        """편집창 내용을 주어진 보고서 데이터로 다시 쓴다."""
        self.report_text.delete("1.0", tk.END)
        self.report_text.insert(
            "1.0",
            format_report_text(values)
        )

    def _choose_report(self):
        """완료보고서 Excel을 선택해 내용을 읽고 편집창에 채운다."""
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="해석 완료보고서 선택",
            filetypes=[("Excel 보고서", "*.xlsm *.xlsx")]
        )
        if not selected:
            return

        try:
            report_data = read_report_data(selected)
        except Exception as error:
            messagebox.showerror(
                "보고서 읽기 실패",
                str(error),
                parent=self.root
            )
            return

        self.report_file = selected
        self.report_summary.set(os.path.basename(selected))
        self._set_report_text(report_data)
        self._append_log("보고서 정보를 불러왔습니다.")
        self._refresh_generate_button()

    def _choose_output_path(self):
        """완성된 PPT를 저장할 파일명과 경로를 대화상자로 선택한다."""
        current = Path(
            self.output_path.get().strip() or str(config.OUTPUT_PPT)
        )
        selected = filedialog.asksaveasfilename(
            parent=self.root,
            title="PPT 저장 위치 선택",
            defaultextension=".pptx",
            filetypes=[("PowerPoint 파일", "*.pptx")],
            initialdir=str(current.parent),
            initialfile=current.name
        )
        if selected:
            self.output_path.set(selected)

    def _choose_images(self):
        """결과 이미지를 추가하고 파일명이 규칙에 맞으면 자동 배정한다."""
        selected = filedialog.askopenfilenames(
            parent=self.root,
            title="해석 결과 이미지 선택",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp")]
        )
        if not selected:
            return

        self.image_files = tuple(dict.fromkeys((*self.image_files, *selected)))
        auto_count = 0
        # 이름이 바뀐 파일은 파일명만으로 Case·부품·Max/Min을 복원한다.
        for image_file in selected:
            metadata = parse_image_metadata(image_file)
            if not metadata["case"] or not metadata["part"]:
                continue

            case_name = metadata["case"]
            part = metadata["part"]
            self.case_assignments.setdefault(case_name, {})
            previous_image = self.case_assignments[case_name].get(part)
            if previous_image and previous_image != image_file:
                self.image_roles.pop(previous_image, None)
            self.case_assignments[case_name][part] = image_file
            self.image_roles[image_file] = (case_name, part)
            if metadata["max"] or metadata["min"]:
                self.legend_values_by_image[image_file] = {
                    "max": metadata["max"],
                    "min": metadata["min"]
                }
            if case_name not in self.case_order:
                self.case_order.append(case_name)
            auto_count += 1

        self._refresh_case_order()
        self._refresh_image_list()
        self.image_summary.set(
            f"이미지 {len(self.image_files)}개 선택됨"
        )
        self._append_log(
            f"이미지 {len(selected)}개 추가, 파일명 자동 분류 {auto_count}개"
        )
        self._refresh_generate_button()

    def _refresh_image_list(self, selected_file=None):
        """필터 조건에 맞는 이미지 목록을 다시 그리고 선택을 유지한다."""
        if selected_file is None:
            current_selection = self.image_list.curselection()
            if current_selection and current_selection[0] < len(self.visible_image_files):
                selected_file = self.visible_image_files[current_selection[0]]

        part_filter = self.image_filter.get()

        def matches_filter(image_file):
            assignment = self.image_roles.get(image_file)
            if part_filter == "전체":
                return True
            if part_filter == "미지정":
                return assignment is None
            return assignment is not None and assignment[1] == part_filter.lower()

        self.visible_image_files = tuple(
            image_file
            for image_file in self.image_files
            if matches_filter(image_file)
        )

        self.image_list.delete(0, tk.END)
        for image_file in self.visible_image_files:
            self.image_list.insert(
                tk.END,
                self._image_display_name(image_file)
            )

        if not self.visible_image_files:
            self.preview_photo = None
            self.preview_original = None
            self.current_preview_file = None
            self.preview_info.set("목록에서 이미지를 선택하면 범례와 모델을 추출합니다.")
            for canvas in self.preview_canvases.values():
                canvas.delete("all")
                canvas.create_text(
                    max(canvas.winfo_width(), 1) / 2,
                    max(canvas.winfo_height(), 1) / 2,
                    text="이미지 선택",
                    fill="#667680"
                )
            self.preview_info.set("목록에서 이미지를 선택하면 미리보기가 표시됩니다.")
            self.assign_glass_button.configure(state="disabled")
            self.assign_rack_button.configure(state="disabled")
            return

        if selected_file in self.visible_image_files:
            selected_index = self.visible_image_files.index(selected_file)
        else:
            selected_index = 0

        self.image_list.selection_set(selected_index)
        self.image_list.activate(selected_index)
        self._show_preview()

    def _image_display_name(self, image_file):
        """목록에 표시할 이름을 만든다. 배정·Max/Min 정보를 괄호로 붙인다."""
        metadata = parse_image_metadata(image_file)
        assignment = self.image_roles.get(image_file)
        values = self.legend_values_by_image.get(image_file, {})
        case_name = assignment[0] if assignment else metadata["case"]
        part = assignment[1] if assignment else metadata["part"]
        details = [
            value
            for value in (case_name, part.upper() if part else "")
            if value
        ]
        if values.get("max"):
            details.append(f"Max {values['max']}")
        if values.get("min"):
            details.append(f"Min {values['min']}")
        suffix = f"  [{ ' · '.join(details) }]" if details else ""
        return os.path.basename(image_file) + suffix

    def _show_preview(self, _event=None):
        """선택한 이미지의 범례·모델 미리보기와 Max/Min 입력값을 표시한다."""
        selection = self.image_list.curselection()
        if not selection:
            return

        image_file = self.visible_image_files[selection[0]]
        self._save_legend_values()
        self.current_preview_file = image_file
        self._load_legend_values(image_file)
        assignment = self.image_roles.get(image_file)
        role_description = (
            f"{assignment[0]} · {assignment[1].upper()}"
            if assignment else "미지정"
        )
        self.preview_info.set(f"{os.path.basename(image_file)}  ·  {role_description}")
        button_state = "normal" if self.case_order else "disabled"
        self.assign_glass_button.configure(state=button_state)
        self.assign_rack_button.configure(state=button_state)

        # 미리보기는 이미지당 한 번만 추출하고 이후에는 캐시를 재사용한다.
        if image_file not in self.preview_results:
            preview_dir = os.path.join(
                self.preview_temp_dir.name,
                str(len(self.preview_results))
            )
            os.makedirs(preview_dir, exist_ok=True)
            self.status.set("모델과 범례를 추출하는 중입니다.")
            try:
                self.preview_results[image_file] = ImageProcessor(
                    preview_dir
                ).process(image_file) or {}
            except Exception as error:
                self.preview_results[image_file] = {}
                self._append_log(f"미리보기 처리 오류: {error}")
            finally:
                self.status.set("이미지 미리보기 준비 완료")

        self._render_preview("legend")
        self._render_preview("model")

    def _load_legend_values(self, image_file):
        """해당 이미지에 저장된 Max/Min 값을 입력란에 채운다."""
        values = self.legend_values_by_image.get(image_file, {})
        self.loading_legend_values = True
        try:
            self.max_value.set(values.get("max", ""))
            self.min_value.set(values.get("min", ""))
        finally:
            self.loading_legend_values = False

    def _save_legend_values(self, *_args):
        """입력란의 Max/Min 값을 현재 선택 이미지에 저장한다."""
        # Max/Min 값은 Case 배정 전까지 선택된 원본 이미지 기준으로 저장한다.
        if self.loading_legend_values or not self.current_preview_file:
            return

        self.legend_values_by_image[self.current_preview_file] = {
            "max": self.max_value.get().strip(),
            "min": self.min_value.get().strip()
        }
        if self.current_preview_file in self.visible_image_files:
            index = self.visible_image_files.index(
                self.current_preview_file
            )
            self.image_list.delete(index)
            self.image_list.insert(
                index,
                self._image_display_name(self.current_preview_file)
            )
            self.image_list.selection_set(index)

    def _collect_legend_values(self):
        """모든 배정된 이미지의 Max/Min 값을 Case/부품 구조로 모은다."""
        self._save_legend_values()
        values = {}
        for image_file, (case_name, part) in self.image_roles.items():
            values.setdefault(case_name, {})[part] = (
                self.legend_values_by_image.get(image_file, {})
            )
        return values

    def _render_preview(self, preview_kind):
        """캔버스 크기에 맞춰 추출된 범례/모델 이미지를 그린다."""
        canvas = self.preview_canvases[preview_kind]
        result = self.preview_results.get(self.current_preview_file, {})
        image_path = result.get(f"{preview_kind}_file") if result else None

        canvas.delete("all")
        if not image_path or not os.path.isfile(image_path):
            canvas.create_text(
                max(canvas.winfo_width(), 1) / 2,
                max(canvas.winfo_height(), 1) / 2,
                text="추출 결과 없음",
                fill="#667680"
            )
            self.preview_photos.pop(preview_kind, None)
            return

        try:
            with Image.open(image_path) as image:
                preview = image.copy()
            # 캔버스 크기에 꽉 차게 축소해 범례·모델 모두 크게 보이게 한다.
            preview.thumbnail(
                (
                    max(round(canvas.winfo_width() - 12), 1),
                    max(round(canvas.winfo_height() - 12), 1)
                ),
                Image.Resampling.LANCZOS
            )
            photo = ImageTk.PhotoImage(preview)
        except Exception as error:
            canvas.create_text(
                max(canvas.winfo_width(), 1) / 2,
                max(canvas.winfo_height(), 1) / 2,
                text="미리보기를 열 수 없습니다.",
                fill="#667680"
            )
            self._append_log(f"미리보기 표시 오류: {error}")
            return

        self.preview_photos[preview_kind] = photo
        canvas.create_image(
            canvas.winfo_width() / 2,
            canvas.winfo_height() / 2,
            image=photo,
            anchor="center"
        )

    def _close(self):
        """미리보기 임시 폴더를 정리하고 창을 닫는다."""
        self.preview_temp_dir.cleanup()
        self.root.destroy()

    def _refresh_case_order(self, selected_case=None):
        """Case 목록을 다시 그리고 Glass/Rack 배정 상태를 함께 표시한다."""
        self.case_order = [
            case_name
            for case_name in self.case_order
            if case_name in self.case_assignments
        ]
        self.case_order.extend(
            case_name
            for case_name in self.case_assignments
            if case_name not in self.case_order
        )

        self.case_list.delete(0, tk.END)
        for case_name in self.case_order:
            parts = self.case_assignments[case_name]
            glass_status = "Glass ✓" if parts.get("glass") else "Glass —"
            rack_status = "Rack ✓" if parts.get("rack") else "Rack —"
            self.case_list.insert(
                tk.END,
                f"{case_name}   {glass_status}   {rack_status}"
            )

        if self.case_order:
            case_name = selected_case if selected_case in self.case_order else self.case_order[0]
            index = self.case_order.index(case_name)
            self.case_list.selection_set(index)
            self.case_list.activate(index)
        self._refresh_generate_button()

    def _add_case(self):
        """입력한 이름의 Case를 추가한다. 비어 있으면 자동 번호를 붙인다."""
        case_name = self.case_name.get().strip()
        if not case_name:
            case_name = f"Case {len(self.case_assignments) + 1}"

        if case_name in self.case_assignments:
            messagebox.showwarning(
                "중복 Case",
                f"이미 등록된 Case입니다: {case_name}",
                parent=self.root
            )
            return

        self.case_assignments[case_name] = {}
        self.case_order.append(case_name)
        self._refresh_case_order(case_name)
        self.case_name.set(f"Case {len(self.case_assignments) + 1}")
        self._append_log(f"{case_name}를 추가했습니다.")

    def _assign_current_image(self, part):
        """선택한 이미지를 선택한 Case의 glass/rack 칸에 배정한다."""
        image_selection = self.image_list.curselection()
        case_selection = self.case_list.curselection()
        if not image_selection:
            messagebox.showinfo(
                "이미지 선택",
                "먼저 미리보기 목록에서 이미지를 선택해 주세요.",
                parent=self.root
            )
            return
        if not case_selection:
            messagebox.showinfo(
                "Case 선택",
                "이미지를 배정할 Case를 선택하거나 추가해 주세요.",
                parent=self.root
            )
            return

        image_file = self.visible_image_files[image_selection[0]]
        case_name = self.case_order[case_selection[0]]

        # 같은 이미지가 다른 Case에 배정되어 있었다면 먼저 해제한다.
        previous_assignment = self.image_roles.get(image_file)
        if previous_assignment:
            previous_case, previous_part = previous_assignment
            self.case_assignments[previous_case].pop(previous_part, None)

        # 같은 칸에 이미 배정된 이미지가 있으면 새 이미지로 교체한다.
        previous_image = self.case_assignments[case_name].get(part)
        if previous_image:
            self.image_roles.pop(previous_image, None)

        self.case_assignments[case_name][part] = image_file
        self.image_roles[image_file] = (case_name, part)
        self._refresh_case_order(case_name)
        self._refresh_image_list(image_file)
        self._append_log(
            f"{os.path.basename(image_file)} → {case_name} / {part.upper()}"
        )

    def _move_case(self, direction):
        """선택한 Case를 목록에서 한 칸 위/아래로 옮긴다."""
        selection = self.case_list.curselection()
        if not selection:
            return

        current_index = selection[0]
        target_index = current_index + direction
        if not 0 <= target_index < len(self.case_order):
            return

        self.case_order[current_index], self.case_order[target_index] = (
            self.case_order[target_index],
            self.case_order[current_index]
        )
        selected_case = self.case_order[target_index]
        self._refresh_case_order(selected_case)

    def _refresh_generate_button(self):
        """이미지와 Case 배정이 준비됐을 때만 생성 버튼을 활성화한다."""
        ready = bool(
            self.image_files
            and build_cases_from_assignments(
                self.case_assignments
            )
        )
        self.generate_button.configure(
            state="normal" if ready and not self.is_busy else "disabled"
        )

    def _prepare_generation_images(self):
        """생성에 넘길 배정 맵 복사본과 사용할 이미지 경로 목록을 만든다."""
        assignments = {
            case_name: dict(parts)
            for case_name, parts in self.case_assignments.items()
        }
        image_files = list(dict.fromkeys(
            image_file
            for parts in assignments.values()
            for image_file in parts.values()
        ))
        return assignments, image_files

    def _start_generation(self):
        """입력을 잠그고 현재 설정을 스냅샷으로 저장한 뒤 생성 스레드를 시작한다."""
        if self.is_busy:
            return

        # 저장 경로가 비어 있으면 기본 출력 경로를 사용한다.
        output_text = self.output_path.get().strip()
        output_path = Path(output_text) if output_text else config.OUTPUT_PPT
        self.output_path.set(str(output_path))

        if output_path.exists() and not messagebox.askyesno(
            "파일 덮어쓰기",
            f"기존 파일을 덮어쓸까요?\n{output_path}",
            parent=self.root
        ):
            return

        try:
            (
                self.generation_case_assignments,
                self.generation_images
            ) = self._prepare_generation_images()
        except Exception as error:
            messagebox.showerror(
                "이미지 이름 저장 실패",
                str(error),
                parent=self.root
            )
            return

        self.is_busy = True
        self.generation_case_order = tuple(self.case_order)
        self.generation_report_overrides = parse_report_text(
            self.report_text.get("1.0", "end-1c")
        )
        self.generation_legend_values = build_legend_values(
            self.image_roles,
            self.legend_values_by_image
        )
        self.generation_rename_originals = self.rename_original_files.get()
        self.generation_output_path = output_path
        self.status.set("PPT를 생성하고 있습니다.")
        self._set_input_state("disabled")
        self.generate_button.configure(state="disabled")
        self._append_log("PPT 생성을 시작합니다.")

        worker = threading.Thread(
            target=self._generate_in_background,
            daemon=True
        )
        worker.start()

    def _generate_in_background(self):
        """백그라운드에서 PPT를 생성하고 결과를 이벤트 큐로 전달한다."""
        try:
            output_path, case_count = generate_presentation(
                self.report_file,
                self.generation_images,
                progress_callback=lambda message: self.events.put(
                    ("progress", message)
                ),
                case_order=self.generation_case_order,
                report_overrides=self.generation_report_overrides,
                case_assignments=self.generation_case_assignments,
                legend_values=self.generation_legend_values,
                output_path=self.generation_output_path
            )
            # 저장에 성공한 뒤에만 원본 이미지 파일명을 바꾼다.
            renamed_files = {}
            rename_errors = []
            if self.generation_rename_originals:
                renamed_files, rename_errors = rename_assigned_images(
                    self.generation_case_assignments,
                    self.generation_legend_values
                )
            self.events.put(
                (
                    "success",
                    str(output_path),
                    case_count,
                    renamed_files,
                    rename_errors
                )
            )
        except Exception as error:
            self.events.put(("error", str(error)))

    def _poll_events(self):
        """이벤트 큐를 주기적으로 확인해 진행 상황·완료·오류를 화면에 반영한다."""
        # 백그라운드 스레드는 큐로만 소통하고, 위젯 갱신은 이 Tk 콜백에서만 한다.
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    self.status.set(event[1])
                    self._append_log(event[1])
                elif event[0] == "success":
                    self._finish_generation()
                    output_path, case_count, renamed_files, rename_errors = event[1:]
                    self._apply_renamed_paths(renamed_files)
                    self.status.set("PPT 생성이 완료되었습니다.")
                    self._append_log(f"저장 완료: {output_path}")
                    try:
                        os.startfile(output_path)
                    except (AttributeError, OSError) as error:
                        self._append_log(f"PPT 자동 열기 실패: {error}")
                    messagebox.showinfo(
                        "생성 완료",
                        self._completion_message(
                            output_path,
                            case_count,
                            renamed_files,
                            rename_errors
                        ),
                        parent=self.root
                    )
                elif event[0] == "error":
                    self._finish_generation()
                    self.status.set("생성 중 오류가 발생했습니다.")
                    self._append_log(f"오류: {event[1]}")
                    messagebox.showerror(
                        "PPT 생성 실패",
                        event[1],
                        parent=self.root
                    )
        except queue.Empty:
            pass

        self.root.after(100, self._poll_events)

    def _apply_renamed_paths(self, renamed_files):

        """이름이 바뀐 원본 이미지의 새 경로를 앱의 모든 상태에 반영한다."""

        if not renamed_files:
            return

        self.image_files = tuple(
            renamed_files.get(image_file, image_file)
            for image_file in self.image_files
        )
        self.visible_image_files = tuple(
            renamed_files.get(image_file, image_file)
            for image_file in self.visible_image_files
        )
        self.legend_values_by_image = {
            renamed_files.get(image_file, image_file): values
            for image_file, values in self.legend_values_by_image.items()
        }
        self.image_roles = {
            renamed_files.get(image_file, image_file): assignment
            for image_file, assignment in self.image_roles.items()
        }
        for case_name, parts in self.case_assignments.items():
            for part, image_file in parts.items():
                parts[part] = renamed_files.get(image_file, image_file)

        for old_path, new_path in renamed_files.items():
            if old_path in self.preview_results:
                self.preview_results[new_path] = self.preview_results.pop(old_path)
            if self.current_preview_file == old_path:
                self.current_preview_file = new_path

        self._refresh_image_list(self.current_preview_file)
        self._refresh_case_order()

    def _completion_message(
        self,
        output_path,
        case_count,
        renamed_files,
        rename_errors
    ):
        """완료 대화상자에 보여줄 요약 메시지를 만든다."""
        message = f"{case_count}개 케이스 PPT를 생성했습니다.\n{output_path}"
        if renamed_files:
            message += f"\n\n원본 이미지 이름 변경: {len(renamed_files)}개"
        if rename_errors:
            message += "\n\n이름을 바꾸지 못한 파일:\n" + "\n".join(rename_errors)
        return message

    def _finish_generation(self):
        """생성 종료 후 잠갔던 입력을 다시 활성화한다."""
        self.is_busy = False
        self._set_input_state("normal")
        self._refresh_generate_button()

    def _set_input_state(self, state):
        """생성 중 실수로 값을 바꾸지 못하게 입력 위젯을 잠그거나 연다."""
        for widget in self.input_widgets:
            widget.configure(state=state)

    def _append_log(self, message):
        """상태 표시줄에 진행 메시지를 표시한다."""
        self.status.set(message)

    def run(self):
        """Tkinter 이벤트 루프를 시작한다."""
        self.root.mainloop()
