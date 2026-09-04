"""Korean and English AWS UI labels, kept together for locale regression tests."""

import re


def names(*labels: str) -> re.Pattern[str]:
    return re.compile(r"^(?:" + "|".join(labels) + r")\s*$", re.IGNORECASE)


TRAINING_ASSIGN = names(r"교육\s*할당", r"Assign\s+training")
ASSIGN_TO_USER = names(r"사용자에 할당", r"사용자에게 할당", r"Assign to users?")
USER_DIALOG = names(r"사용자 선택", r"Select users?")
USER_SEARCH = names(r"사용자 찾기", r"사용자 검색", r"Find users?", r"Search users?")
USERS = names(r"사용자", r"Users?")
ASSIGN = names(r"할당", r"Assign")
CANCEL = names(r"취소", r"Cancel")
DONE = names(r"완료", r"Done", r"Complete")
REGISTER_ALL = names(
    r"선택한 모든 사용자를 등록하고 싶습니다\.?",
    r"I (?:want|would like) to (?:enroll|register) all (?:the )?selected users\.?",
    r"(?:Enroll|Register) all selected users\.?",
)
NO_MATCHES = names(
    r"일치 항목 없음",
    r"데이터를 찾을 수 없음",
    r"No matches",
    r"No matching (?:users|results)(?: found)?",
    r"No (?:users|results|data)(?: found)?",
)
LOADING = names(r"로딩(?: 중)?", r"불러오는 중", r"Loading(?:\.\.\.)?")
NEXT_PAGE = names(r"다음(?: 페이지)?", r"Next(?: page)?", r"Go to next page")
ENROLLMENT_HEADER = names(
    r"등록 상태", r"수강 등록 상태", r"Enrollment status", r"Registration status"
)
ENROLLED = names(
    r"등록됨",
    r"등록 완료",
    r"수강 등록됨",
    r"Enrolled",
    r"Registered",
    r"Enrollment complete",
    r"Registration complete",
)
