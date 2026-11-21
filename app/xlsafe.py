"""
Excel 수식 주입 방지

보고서에는 Amazon에서 받은 글자(고객 검색어, 상품명, 리뷰 등)가 들어갑니다. 누군가 ‘=HYPERLINK(...)’ 같은 글자로
검색하면 그대로 셀에 들어가 Excel이 수식으로 실행할 수 있습니다(openpyxl은 ‘=’로 시작하는 글자를 수식으로 저장).
이 모듈을 불러오면, 이 프로그램이 직접 만드는 단순 계산식(셀 참조 + IF/OR/SUM 등)만 수식으로 두고
나머지 ‘=’로 시작하는 글자는 모두 일반 글자로 저장합니다. 이미 있는 Excel 파일을 열어 다시 저장할 때 그 파일의 수식은 그대로 둡니다.
"""

import re

try:
    from openpyxl.cell.cell import Cell
except ImportError:      # openpyxl이 아직 설치되지 않았으면(설치 안내는 각 화면에서) 아무것도 안 함
    Cell = None

_SAFE = re.compile(r'^=(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|IFERROR|IF|OR|AND|SUM|ROUND|ABS|MAX|MIN|AVERAGE|""|[0-9.+\-*/(),=<> %])+$')


def safe_formula(text: str) -> bool:
    return bool(_SAFE.match(text))


if Cell is not None and not getattr(Cell, "_xlsafe", False):
    _orig = Cell._bind_value

    def _bind_value(self, value):
        _orig(self, value)
        if self.data_type == "f" and isinstance(self._value, str) and not safe_formula(self._value):
            self.data_type = "s"

    Cell._bind_value = _bind_value
    Cell._xlsafe = True
