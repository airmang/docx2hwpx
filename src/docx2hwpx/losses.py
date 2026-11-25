"""Bookkeeping for Word features the Hangul copy cannot carry as Word had them."""
from __future__ import annotations

from collections import Counter

from .api import Loss

# feature -> (kind, Korean label template with {n})
CATALOG: dict[str, tuple[str, str]] = {
    "textbox": ("changed", "글상자 {n}개를 본문 글로 바꿨습니다"),
    "shape": ("dropped", "도형 {n}개가 빠졌습니다"),
    "chart": ("dropped", "차트 {n}개가 빠졌습니다"),
    "smartart": ("dropped", "SmartArt {n}개가 빠졌습니다"),
    "equation": ("changed", "수식 {n}개를 글자로 바꿨습니다"),
    "comment": ("dropped", "메모 {n}개가 빠졌습니다"),
    "tracked_change": ("changed", "변경 내용 추적 {n}곳을 최종본으로 받아들였습니다"),
    "floating_picture": ("changed", "떠 있는 그림 {n}개를 글자처럼 넣었습니다"),
    "ole_object": ("changed", "OLE 개체 {n}개를 그림으로 바꿨습니다"),
    "ole_dropped": ("dropped", "OLE 개체 {n}개가 빠졌습니다"),
    "image_unsupported": ("dropped", "열 수 없는 그림 {n}개가 빠졌습니다"),
    "linked_image": ("dropped", "바깥 파일에 연결된 그림 {n}개가 빠졌습니다"),
    "field": ("changed", "자동 필드 {n}개를 결과 글자로 바꿨습니다"),
    "internal_link": ("changed", "문서 안 링크 {n}개를 보통 글자로 바꿨습니다"),
    "hidden_text": ("dropped", "숨긴 글자 {n}곳이 빠졌습니다"),
    "first_page_header": ("changed", "첫 쪽만 다른 머리말·꼬리말 {n}개를 모든 쪽에 같게 넣었습니다"),
    "alt_chunk": ("dropped", "끼워 넣은 다른 문서 {n}개가 빠졌습니다"),
    "page_border": ("dropped", "쪽 테두리 {n}개가 빠졌습니다"),
    "line_numbers": ("dropped", "줄 번호 {n}개 구역 설정이 빠졌습니다"),
    "vertical_text": ("changed", "세로 쓰기 {n}곳을 가로 쓰기로 바꿨습니다"),
    "form_control": ("changed", "입력 양식 {n}개를 글자로 바꿨습니다"),
}


class LossLog:
    def __init__(self) -> None:
        self.counts: Counter[str] = Counter()

    def add(self, feature: str, n: int = 1) -> None:
        if n > 0:
            self.counts[feature] += n

    def to_list(self) -> list[Loss]:
        out = []
        for feature, n in sorted(self.counts.items()):
            kind, label = CATALOG.get(feature, ("changed", feature + " {n}곳이 달라졌습니다"))
            out.append(Loss(feature=feature, kind=kind, count=n, label=label.format(n=n)))  # type: ignore[arg-type]
        return out
