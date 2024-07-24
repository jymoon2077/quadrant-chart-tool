DEBUG = False  # 디버그 모드 설정
TEXT_COLUMN_LIST = ["Key", "Summary"]  # 텍스트 값을 갖는 컬럼 정의
ANNOTAION_DEFAULT_SIZE = 10
ANNOTAION_BIG_SIZE = 19


def debug_print(message):
    if DEBUG:
        print(message)
