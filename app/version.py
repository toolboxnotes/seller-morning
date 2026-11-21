"""
프로그램 버전(유의적 버전 MAJOR.MINOR.PATCH)

  MAJOR  예전 설정·자료와 맞지 않게 바뀔 때(설정을 다시 해야 하거나 자료 형식이 바뀜)
  MINOR  새 기능을 더할 때(예전 것은 그대로 동작)
  PATCH  고치기만 할 때(오류 수정·문구·작은 개선)

바꿀 때: 아래 숫자와 RELEASED를 고치고, README 첫 부분의 ‘현재 버전’과 맨 아래 ‘버전 기록’에 한 줄 추가.
"""

MAJOR = 1
MINOR = 2
PATCH = 0
RELEASED = "2026-11-06"

__version__ = f"{MAJOR}.{MINOR}.{PATCH}"
VERSION = __version__


def text() -> str:
    """화면·로그에 쓰는 표기, 예: v1.0.0 (2026-11-05)"""
    return f"v{__version__} ({RELEASED})"
