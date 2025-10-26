import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from typing import List, Tuple, Any, Union, Optional
from openpyxl.worksheet.worksheet import Worksheet
import os
from core.env import EXCEL_DIR
import uuid
import time


class ExcelTool:
    def __init__(
            self,
            default_headers: Optional[List[str]] = None,
            header_row: int = 1,
            file_name: Optional[str] = None
    ):
        if header_row < 1:
            raise ValueError("Header row must be 1 or greater.")
        self.default_headers = default_headers or ['序号', 'BD', '达人名称', '带货店铺', '样品名', '合作类型', '发样数量','性别','粉丝量','履约率','主页链接','佣金率','寄样批准日期','发货日期','订单号','履约方式','创作视频链接','视频码']
        self.header_row = header_row
        if not file_name:
            file_name = f"excel_{uuid.uuid4().hex[:8]}_{int(time.time())}.xlsx"
        self.file_path = EXCEL_DIR.joinpath(file_name)
        self.workbook = self.load_workbook(self.file_path) if os.path.exists(self.file_path) else self.create_workbook()

    def create_workbook(self) -> openpyxl.Workbook:
        return openpyxl.Workbook()

    def load_workbook(self, file_path: str) -> openpyxl.Workbook:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件路径 {file_path} 不存在.")
        try:
            return openpyxl.load_workbook(file_path)
        except Exception as e:
            raise ValueError(f"加载文件失败: {str(e)}")

    def save(self) -> None:
        self.workbook.save(self.file_path)

    def get_sheet_names(self) -> List[str]:
        return self.workbook.sheetnames

    def get_sheet(self, sheet_name: str) -> openpyxl.worksheet.worksheet.Worksheet:
        if sheet_name not in self.workbook.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' 不存在")
        return self.workbook[sheet_name]

    def create_sheet(
            self,
            sheet_name: str,
            headers: Optional[List[str]] = None,
            position: Optional[int] = None,
            header_font: Optional[Font] = None,
            header_alignment: Optional[Alignment] = None,
            header_fill: Optional[PatternFill] = None
    ) -> openpyxl.worksheet.worksheet.Worksheet:
        sheet = self.workbook.create_sheet(sheet_name, position)
        headers_to_write = headers if headers is not None else self.default_headers
        if headers_to_write:
            for col, header in enumerate(headers_to_write, start=1):
                self.write_cell(
                    sheet, self.header_row, col, header,
                    font=header_font, alignment=header_alignment, fill=header_fill
                )
        return sheet

    def set_default_headers(self, headers: List[str]) -> None:
        self.default_headers = headers

    def check_headers(self, sheet: openpyxl.worksheet.worksheet.Worksheet,
                      expected_headers: Optional[List[str]] = None) -> bool:
        if self.header_row > sheet.max_row:
            raise ValueError(f"Header row {self.header_row} 超过了长度限制.")
        expected = expected_headers if expected_headers is not None else self.default_headers
        if not expected:
            return True
        actual_headers = [sheet.cell(row=self.header_row, column=col).value for col in range(1, len(expected) + 1)]
        return actual_headers == expected

    def write_cell(
            self,
            sheet: openpyxl.worksheet.worksheet.Worksheet,
            row: int,
            col: int,
            value: Any,
            font: Optional[Font] = None,
            alignment: Optional[Alignment] = None,
            fill: Optional[PatternFill] = None
    ) -> None:
        cell = sheet.cell(row=row, column=col)
        cell.value = value
        if font:
            cell.font = font
        if alignment:
            cell.alignment = alignment
        if fill:
            cell.fill = fill

    def read_cell(self, sheet: openpyxl.worksheet.worksheet.Worksheet, row: int, col: int) -> Any:
        return sheet.cell(row=row, column=col).value

    def append_row(self, sheet: openpyxl.worksheet.worksheet.Worksheet, data: List[Any]) -> None:
        sheet.append(data)

    def get_row_count(self, sheet: openpyxl.worksheet.worksheet.Worksheet) -> int:
        return sheet.max_row

    def get_column_count(self, sheet: openpyxl.worksheet.worksheet.Worksheet) -> int:
        return sheet.max_column

    def read_range(
            self,
            sheet: openpyxl.worksheet.worksheet.Worksheet,
            start_row: int,
            start_col: int,
            end_row: int,
            end_col: int,
            skip_headers: bool = True
    ) -> List[List[Any]]:
        if skip_headers and start_row <= self.header_row:
            start_row = self.header_row + 1
        return [
            [sheet.cell(row=row, column=col).value for col in range(start_col, end_col + 1)]
            for row in range(start_row, end_row + 1)
        ]

    def write_range(
            self,
            sheet: openpyxl.worksheet.worksheet.Worksheet,
            start_row: int,
            start_col: int,
            data: List[List[Any]]
    ) -> None:
        if start_row <= self.header_row:
            start_row = self.header_row + 1
        for row_idx, row_data in enumerate(data, start=start_row):
            for col_idx, value in enumerate(row_data, start=start_col):
                sheet.cell(row=row_idx, column=col_idx).value = value

    def set_column_width(self, sheet: openpyxl.worksheet.worksheet.Worksheet, column: int, width: float) -> None:
        column_letter = get_column_letter(column)
        sheet.column_dimensions[column_letter].width = width

    def set_row_height(self, sheet: openpyxl.worksheet.worksheet.Worksheet, row: int, height: float) -> None:
        sheet.row_dimensions[row].height = height

    def apply_format_to_range(
            self,
            sheet: openpyxl.worksheet.worksheet.Worksheet,
            start_row: int,
            start_col: int,
            end_row: int,
            end_col: int,
            font: Optional[Font] = None,
            alignment: Optional[Alignment] = None,
            fill: Optional[PatternFill] = None
    ) -> None:
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row=row, column=col)
                if font:
                    cell.font = font
                if alignment:
                    cell.alignment = alignment
                if fill:
                    cell.fill = fill

