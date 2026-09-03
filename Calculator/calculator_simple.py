# -*- coding: utf-8 -*-
"""
calculator_simple.py  v3.0
간이법 산정양식 사양

  2024-간이법 : 1101_SW사업_구현단계_SW개발비_간이법_산정양식 (.xlsx)
                시트 FP산정(간이법) / SW개발비 산정, 2024 대가산정 가이드

간이법은 복잡도를 따지지 않고 유형별 평균 가중치를 씁니다.
  ILF 7.5 / EIF 5.4 / EI 4 / EO 5.2 / EQ 3.9
가중치 값은 계산 엔진(calculator_ui.py)의 SIMPLE_WEIGHT 에 있습니다.

항목의 뜻은 calculator_detail.py 머리말과 같습니다.

실행하면 이 양식만 목록에 띄우고 화면을 연다.
  python calculator_simple.py [파일경로]
"""

FORMS = {
    "2024-간이법": {
        "code": "1101", "group": "구현단계",
        "name": "SW개발비 · 간이법 (2024 가이드)",
        "file": "1101.SW사업 구현단계_SW개발비_간이법_산정양식_241231.xlsx",
        "fp": {
            "sheets": ["FP산정(간이법)"],
            "anchor": {"app": (2, ["애플리케이션명", "어플리케이션명"]),
                       "type": (6, ["FP유형"])},
            "cols": [
                {"k": "app", "x": 2, "t": "text", "label": "①어플리케이션명", "w": 130},
                {"k": "biz", "x": 3, "t": "text", "label": "②세부 업무명", "w": 120},
                {"k": "proc", "x": 4, "t": "text", "label": "③단위프로세스명", "w": 210},
                {"k": "desc", "x": 5, "t": "text", "label": "단위프로세스 설명", "w": 300},
                {"k": "type", "x": 6, "t": "sel", "label": "④FP유형", "w": 80},
                {"k": "wt", "x": 7, "t": "ro", "label": "⑤가중치", "w": 70},
                {"k": "remark", "x": 8, "t": "text", "label": "비고", "w": 180},
            ],
            "method": "간이법",
        },
        "cost": {
            "sheets": ["SW개발비 산정", "SW개발비산정"],
            "kind": "kosa2024",
            "check": {"B5": "총기능점수", "B21": "SW규모"},
        },
    },
}


if __name__ == "__main__":
    # 2016 양식은 한 파일로 간이법·상세법을 겸하므로 여기서도 고를 수 있게 함께 띄운다.
    import calculator_ui
    calculator_ui.main(only=list(FORMS) + ["2016-기능점수"])