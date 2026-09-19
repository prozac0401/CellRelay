"""AWS Skill Builder-specific user-assignment workflow.

This module deliberately contains every site-specific locator so the generic
Playwright worker and the UI remain reusable for other pages.
"""

from __future__ import annotations

import asyncio
import re
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass
from urllib.parse import parse_qs, urlparse

from playwright.async_api import Locator, Page

from app.browser import aws_labels as labels
from app.browser.search_observer import SearchObserver

_TRAINING_ASSIGN_BUTTON = labels.TRAINING_ASSIGN
_ASSIGN_TO_USER_ITEM = labels.ASSIGN_TO_USER
_USER_DIALOG = labels.USER_DIALOG
_USER_SEARCH = labels.USER_SEARCH
_USER_TABLE = labels.USERS
_ASSIGN_BUTTON = labels.ASSIGN
_CANCEL_BUTTON = labels.CANCEL
_CONFIRM_ASSIGN_BUTTON = labels.DONE
_REGISTER_SELECTED_USERS = labels.REGISTER_ALL
_ASSIGNMENT_CONFIRMATION_TEST_ID = "assign_classroom_training_confirmation_modal"


class AwsAssignmentError(RuntimeError):
    """Base error for a safe, stopped AWS assignment attempt."""


class AwsAssignmentCancelled(AwsAssignmentError):
    """Raised when Pause/Stop control cancels the current operation."""


class AwsUserLookupError(AwsAssignmentError):
    """Raised unless the search produces exactly one matching user row."""


@dataclass(frozen=True, slots=True)
class AwsAssignmentResult:
    """UI workflow outcome, not independent verification of AWS enrollment.

    ``assigned`` means Done was clicked with enrollment checked and its modal
    closed. A short settling delay follows unless Stop is requested. The
    administrator audits the roster.
    """

    assigned: bool
    matched_user_label: str = ""
    skip_reason: str = ""


class AwsSkillBuilderAssignment:
    """Assign one user on the observed AWS Skill Builder training-detail page."""

    def __init__(
        self,
        page: Page,
        stop_event: threading.Event,
        pause_event: threading.Event,
        on_stage: Callable[[str, str], None] | None = None,
        *,
        lookup_timeout_ms: int = 15_000,
        result_stable_ms: int = 500,
        invalid_result_stable_ms: int = 3_000,
        readiness_timeout_ms: int = 20_000,
        post_confirmation_settle_ms: int = 1_500,
        post_cancel_settle_ms: int = 750,
        enforce_aws_host: bool = True,
    ) -> None:
        self._page = page
        self._stop_event = stop_event
        self._pause_event = pause_event
        self._on_stage = on_stage or (lambda _stage, _message: None)
        self._lookup_timeout_ms = max(1_000, lookup_timeout_ms)
        self._result_stable_ms = max(300, result_stable_ms)
        self._invalid_result_stable_ms = max(
            self._result_stable_ms,
            invalid_result_stable_ms,
        )
        self._readiness_timeout_ms = max(1_000, readiness_timeout_ms)
        self._post_confirmation_settle_ms = max(0, post_confirmation_settle_ms)
        self._post_cancel_settle_ms = max(0, post_cancel_settle_ms)
        self._enforce_aws_host = enforce_aws_host
        self._paused_seconds = 0.0
        self._expected_url: str | None = None

    async def validate_ready(self, expected_url: str | None = None) -> str:
        """Verify the intended training page without clicking or entering data."""
        self._validate_target_page(expected_url)
        self._expected_url = expected_url or self._page.url
        await self._checkpoint()
        await self._wait_for_training_assign_button()
        return self._page.url

    async def assign_user(
        self,
        search_value: str,
        expected_url: str | None = None,
    ) -> AwsAssignmentResult:
        """Run search -> matching row -> assign -> checked Done -> modal closed."""
        query = search_value.strip()
        if not query:
            raise AwsUserLookupError("빈 Excel 값은 사용자 검색에 사용할 수 없습니다.")
        self._validate_target_page(expected_url)
        self._expected_url = expected_url or self._page.url
        await self._checkpoint()

        self._stage("OPENING_ASSIGNMENT", "교육 할당 메뉴를 여는 중입니다.")
        dialog = await self._open_user_dialog()
        await self._checkpoint()

        self._stage(
            "SEARCHING_USER",
            "사용자 찾기에 현재 셀 값을 입력하고 검색을 실행하는 중입니다.",
        )
        search = await self._find_user_search_input(dialog)
        table = await self._single(
            dialog.get_by_role("table", name=_USER_TABLE, exact=True),
            "사용자 검색 결과 표",
        )
        baseline_signature = await self._row_signature(self._data_rows(table))
        await search.fill(query, timeout=10_000)
        if (await search.input_value(timeout=5_000)).strip() != query:
            raise AwsAssignmentError("사용자 찾기 입력값을 확인하지 못했습니다.")
        observer = SearchObserver(self._page, query)
        observer.start()
        try:
            await search.press("Enter", timeout=5_000)
            row = await self._wait_for_one_matching_row(
                table,
                query,
                baseline_signature,
                observer,
            )
        except AwsUserLookupError as exc:
            return await self._cancel_unmatched_user(dialog, str(exc))
        finally:
            observer.close()
        self._stage("SELECTING_USER", "검색 결과 1개를 확인하고 선택했습니다.")
        checkbox = await self._single(
            row.get_by_role("checkbox"),
            "검색 결과 체크박스",
        )
        await checkbox.check(timeout=5_000)
        if not await checkbox.is_checked():
            raise AwsAssignmentError("검색 결과 체크박스가 선택되지 않았습니다.")
        matched_label = await self._matched_user_label(row, query)

        assign_button = dialog.get_by_test_id(
            "assign-classroom-training-assign-modal-btn"
        )
        if await assign_button.count() != 1:
            assign_button = await self._single(
                dialog.get_by_role("button", name=_ASSIGN_BUTTON, exact=True),
                "할당 버튼",
            )
        await self._wait_until_enabled(assign_button, 5_000)
        await self._checkpoint()
        # Revalidate identity immediately before the external mutation.
        current_email = await self._matched_user_label(row, query)
        if (
            await self._data_rows(table).count() != 1
            or not self.matches_user(query, current_email)
            or self._normalize(current_email) != self._normalize(matched_label)
        ):
            raise AwsAssignmentError(
                "SEARCH_CHANGED: 선택 후 검색 결과가 변경되었습니다."
            )
        if (
            await self._data_rows(table).get_by_role("checkbox", checked=True).count()
            != 1
        ):
            raise AwsAssignmentError(
                "SELECTION_CHANGED: 선택된 사용자가 정확히 1명이 아닙니다."
            )

        self._stage("ASSIGNING_USER", "선택한 사용자를 교육에 할당하는 중입니다.")
        await assign_button.click(timeout=10_000)
        await self._wait_until_hidden(dialog, 15_000)

        self._stage(
            "CONFIRMING_ASSIGNMENT",
            "AWS 사용자 등록 확인 창에서 체크박스를 선택하고 완료하는 중입니다.",
        )
        await self._confirm_assignment()

        self._stage(
            "ASSIGNMENT_SUBMITTED",
            "완료 버튼 처리와 팝업 종료를 확인했습니다. 최종 명단은 관리자가 대조합니다.",
        )
        # Record the completed UI submission, not a roster-verified enrollment.
        # Never inspect the main table: it may show another page or stale data.
        # Next-row button readiness is checked only after this receipt is saved.
        return AwsAssignmentResult(
            assigned=True,
            matched_user_label=matched_label,
        )

    def _validate_target_page(self, expected_url: str | None = None) -> None:
        if not self._enforce_aws_host:
            return
        if not self.is_training_page(self._page.url):
            raise AwsAssignmentError(
                "AWS Skill Builder 교육 상세 페이지가 아닙니다. URL을 확인하세요."
            )
        if expected_url and not self.is_same_training_destination(
            expected_url,
            self._page.url,
        ):
            raise AwsAssignmentError(
                "현재 브라우저가 입력한 교육 상세 URL과 다른 페이지에 있습니다. "
                "대상 교육과 조직을 확인하세요."
            )

    async def _open_user_dialog(self) -> Locator:
        dialog = self._page.get_by_role("dialog", name=_USER_DIALOG, exact=True)
        if await dialog.count() == 1 and await dialog.is_visible():
            # A dialog left over from Stop/Error can retain filter chips and
            # checked users. Reopen it so a new query starts from a clean state.
            cancel = await self._single_visible(
                dialog.get_by_role("button", name=labels.CANCEL),
                "사용자 선택 취소 버튼",
            )
            await cancel.click(timeout=5_000)
            await self._wait_until_hidden(dialog, 10_000)

        # A new assignment must start from a newly resolved button. AWS can
        # rerender or briefly remove the control after the previous modal closes.
        assign_menu_button = await self._wait_for_training_assign_button()
        await assign_menu_button.click(timeout=10_000)
        await self._checkpoint()
        assign_to_user = await self._wait_for_single_visible(
            self._page.get_by_role(
                "menuitem",
                name=_ASSIGN_TO_USER_ITEM,
                exact=True,
            ),
            "사용자에 할당 메뉴",
            self._readiness_timeout_ms,
        )
        await assign_to_user.click(timeout=10_000)
        await self._wait_until_visible(dialog, 10_000)
        return await self._single(dialog, "사용자 선택 창")

    async def _wait_for_training_assign_button(self) -> Locator:
        """Reacquire the training action after AWS returns to the detail page."""
        button = await self._wait_for_single_visible(
            self._page.get_by_role(
                "button",
                name=_TRAINING_ASSIGN_BUTTON,
            ),
            "교육 할당 버튼",
            self._readiness_timeout_ms,
        )
        await self._wait_until_enabled(button, self._readiness_timeout_ms)
        return button

    async def _cancel_unmatched_user(
        self,
        dialog: Locator,
        reason: str,
    ) -> AwsAssignmentResult:
        """Close a pre-assignment lookup failure and restore the detail page."""
        self._stage(
            "CANCELLING_UNMATCHED_USER",
            "사용자를 추가하지 않고 사용자 선택 창을 취소하는 중입니다.",
        )
        cancel_button = await self._single_visible(
            dialog.get_by_role(
                "button",
                name=_CANCEL_BUTTON,
                exact=True,
            ),
            "사용자 선택 취소 버튼",
        )
        await self._wait_until_enabled(cancel_button, 5_000)
        await self._checkpoint()
        await cancel_button.click(timeout=10_000)
        await self._wait_until_hidden(
            dialog,
            15_000,
            "사용자 선택 취소 후 창이 닫히지 않았습니다.",
        )
        if self._post_cancel_settle_ms:
            await self._interruptible_wait(self._post_cancel_settle_ms / 1000)
        await self._checkpoint()
        return AwsAssignmentResult(
            assigned=False,
            skip_reason=reason,
        )

    async def _find_user_search_input(self, dialog: Locator) -> Locator:
        """Wait for AWS input rendered inside or alongside the portaled dialog."""
        candidates = (
            dialog.get_by_placeholder(_USER_SEARCH),
            dialog.get_by_role("searchbox", name=_USER_SEARCH),
            dialog.get_by_role("textbox", name=_USER_SEARCH),
            dialog.get_by_role("combobox", name=_USER_SEARCH),
            self._page.get_by_placeholder(_USER_SEARCH),
            self._page.get_by_role("searchbox", name=_USER_SEARCH),
            self._page.get_by_role("textbox", name=_USER_SEARCH),
            self._page.get_by_role("combobox", name=_USER_SEARCH),
        )
        deadline = self._now() + self._readiness_timeout_ms / 1000
        while self._now() < deadline:
            await self._checkpoint()
            for candidate in candidates:
                editable = [
                    item
                    for item in await candidate.all()
                    if await item.is_visible() and await item.is_editable()
                ]
                if len(editable) == 1:
                    return editable[0]
                if len(editable) > 1:
                    raise AwsAssignmentError(
                        "화면에 보이고 입력 가능한 사용자 찾기 영역이 "
                        f"정확히 1개여야 합니다. 현재 {len(editable)}개입니다."
                    )
            await self._interruptible_wait(0.2)
        raise AwsAssignmentError(
            "사용자 찾기 입력 영역이 "
            f"{self._readiness_timeout_ms / 1000:g}초 안에 준비되지 않았습니다. "
            "사용자 선택 창이 완전히 열린 상태인지 확인하세요."
        )

    async def _confirm_assignment(self) -> None:
        """Opt in to registration and complete the final AWS confirmation."""
        modal = self._page.get_by_test_id(_ASSIGNMENT_CONFIRMATION_TEST_ID)
        await self._wait_until_visible(
            modal,
            15_000,
            "할당 확인 창이 열리지 않았습니다.",
        )
        modal = await self._single_visible(modal, "할당 확인 창")

        # AWS currently renders a single checkbox in this modal. Prefer its
        # accessible label, but retain a guarded fallback for UI libraries that
        # visually associate the label without exposing that association.
        named_checkboxes = modal.get_by_role(
            "checkbox",
            name=_REGISTER_SELECTED_USERS,
        )
        visible_named = [
            item for item in await named_checkboxes.all() if await item.is_visible()
        ]
        if len(visible_named) == 1:
            registration_checkbox = visible_named[0]
        elif len(visible_named) > 1:
            raise AwsAssignmentError(
                "사용자 등록 확인 체크박스가 정확히 1개여야 합니다. "
                f"현재 {len(visible_named)}개입니다."
            )
        else:
            prompt = await self._single_visible(
                modal.get_by_text(_REGISTER_SELECTED_USERS, exact=True),
                "사용자 등록 확인 문구",
            )
            if not await prompt.is_visible():
                raise AwsAssignmentError("사용자 등록 확인 문구가 보이지 않습니다.")
            registration_checkbox = await self._single_visible(
                modal.get_by_role("checkbox"),
                "사용자 등록 확인 체크박스",
            )

        await self._checkpoint()
        await registration_checkbox.check(timeout=5_000)
        if not await registration_checkbox.is_checked():
            raise AwsAssignmentError("사용자 등록 확인 체크박스가 선택되지 않았습니다.")

        confirm_button = await self._single_visible(
            modal.get_by_role(
                "button",
                name=_CONFIRM_ASSIGN_BUTTON,
                exact=True,
            ),
            "할당 확인 버튼",
        )
        await self._wait_until_enabled(confirm_button, 5_000)
        await self._checkpoint()
        await confirm_button.click(timeout=10_000)
        await self._wait_until_hidden(
            modal,
            15_000,
            "최종 할당 후 확인 창이 닫히지 않았습니다.",
        )
        await self._settle_closed_confirmation()

    async def _settle_closed_confirmation(self) -> None:
        """Preserve an observed UI submission while delaying the next row."""
        # Done was clicked and its modal was observed closed. Pause/Stop must
        # not turn that receipt into a cancellation and allow a duplicate retry.
        # This records UI submission only; the administrator audits enrollment.
        deadline = time.monotonic() + self._post_confirmation_settle_ms / 1000
        while time.monotonic() < deadline:
            if self._stop_event.is_set():
                return
            await asyncio.sleep(min(0.1, max(0, deadline - time.monotonic())))

    async def _wait_for_one_matching_row(
        self,
        table: Locator,
        query: str,
        baseline_signature: tuple[str, ...],
        observer: SearchObserver,
    ) -> Locator:
        deadline = self._now() + self._lookup_timeout_ms / 1000
        last_signature: tuple[str, ...] | None = None
        stable_since = self._now()
        dialog = self._page.get_by_role("dialog", name=labels.USER_DIALOG)
        saw_busy = False

        while self._now() < deadline:
            await self._checkpoint()
            if observer.failed:
                raise AwsAssignmentError(
                    "SEARCH_FAILED: 사용자 검색 요청이 실패했습니다. 다시 확인하세요."
                )
            busy_elements = dialog.locator('[aria-busy="true"], [role="progressbar"]')
            busy = any([await e.is_visible() for e in await busy_elements.all()])
            busy = busy or any(
                [
                    await e.is_visible()
                    for e in await dialog.get_by_label(labels.LOADING).all()
                ]
            )
            saw_busy = saw_busy or busy
            # A changed row count alone is not evidence that a search finished.
            # Require the submitted request to finish, or a complete UI loading
            # cycle, before classifying even an empty result as a skipped user.
            finished = (
                (observer.completed or saw_busy) and not busy and not observer.pending
            )
            if not finished:
                last_signature = None
                stable_since = self._now()
                await self._interruptible_wait(0.1)
                continue
            if any(
                [await e.is_visible() for e in await dialog.get_by_role("alert").all()]
            ):
                raise AwsAssignmentError(
                    "SEARCH_FAILED: 검색 창에 오류 알림이 표시되었습니다."
                )
            rows = self._data_rows(table)
            signature = await self._row_signature(rows)
            if signature != last_signature:
                last_signature = signature
                stable_since = self._now()

            stable_ms = (self._now() - stable_since) * 1000
            if stable_ms >= self._result_stable_ms:
                next_buttons = dialog.get_by_role("button", name=labels.NEXT_PAGE)
                has_more = any(
                    [
                        await b.is_visible() and await b.is_enabled()
                        for b in await next_buttons.all()
                    ]
                )
                if len(signature) > 1 or has_more:
                    raise AwsUserLookupError(
                        "MULTIPLE_MATCHES: 검색 결과가 여러 명입니다."
                    )
                if len(signature) == 1:
                    email = await self._matched_user_label(rows.first, query)
                    if self.matches_user(query, email):
                        return rows.first
                    raise AwsUserLookupError(
                        "IDENTITY_MISMATCH: 검색 결과의 이메일/사용자 ID가 일치하지 않습니다."
                    )
                empty = dialog.get_by_text(labels.NO_MATCHES, exact=True)
                if any([await e.is_visible() for e in await empty.all()]):
                    raise AwsUserLookupError(
                        "NOT_FOUND: 검색을 완료했지만 일치하는 사용자가 없습니다."
                    )
            await self._interruptible_wait(0.1)

        raise AwsAssignmentError(
            "SEARCH_TIMEOUT: 검색 완료를 확인하지 못했습니다. 현재 셀에서 중지합니다."
        )

    async def _wait_until_enabled(self, locator: Locator, timeout_ms: int) -> None:
        deadline = self._now() + timeout_ms / 1000
        while self._now() < deadline:
            await self._checkpoint()
            if await locator.is_enabled():
                return
            await self._interruptible_wait(0.1)
        raise AwsAssignmentError("할당 버튼이 활성화되지 않았습니다.")

    async def _wait_for_single_visible(
        self,
        locator: Locator,
        description: str,
        timeout_ms: int,
    ) -> Locator:
        """Wait for one React-rendered element while preserving strict matching."""
        deadline = self._now() + timeout_ms / 1000
        while self._now() < deadline:
            await self._checkpoint()
            visible = [item for item in await locator.all() if await item.is_visible()]
            if len(visible) == 1:
                return visible[0]
            if len(visible) > 1:
                raise AwsAssignmentError(
                    f"화면에 보이는 {description}이(가) 정확히 1개여야 합니다. "
                    f"현재 {len(visible)}개입니다."
                )
            await self._interruptible_wait(0.2)
        raise AwsAssignmentError(
            f"{description}이(가) {timeout_ms / 1000:g}초 안에 준비되지 않았습니다."
        )

    async def _wait_until_visible(
        self,
        locator: Locator,
        timeout_ms: int,
        timeout_message: str = "사용자 선택 창이 열리지 않았습니다.",
    ) -> None:
        deadline = self._now() + timeout_ms / 1000
        while self._now() < deadline:
            await self._checkpoint()
            if await locator.count() == 1 and await locator.is_visible():
                return
            await self._interruptible_wait(0.1)
        raise AwsAssignmentError(timeout_message)

    async def _wait_until_hidden(
        self,
        locator: Locator,
        timeout_ms: int,
        timeout_message: str = "할당 후 사용자 선택 창이 닫히지 않았습니다.",
    ) -> None:
        deadline = self._now() + timeout_ms / 1000
        while self._now() < deadline:
            await self._checkpoint()
            if await locator.count() == 0 or not await locator.is_visible():
                return
            await self._interruptible_wait(0.1)
        raise AwsAssignmentError(timeout_message)

    async def _checkpoint(self) -> None:
        pause_started = time.monotonic()
        while self._pause_event.is_set():
            if self._stop_event.is_set():
                raise AwsAssignmentCancelled("AWS 사용자 할당이 중지되었습니다.")
            await asyncio.sleep(0.1)
        self._paused_seconds += time.monotonic() - pause_started
        if self._stop_event.is_set():
            raise AwsAssignmentCancelled("AWS 사용자 할당이 중지되었습니다.")
        if self._page.is_closed():
            raise AwsAssignmentError("브라우저 페이지가 닫혔습니다.")
        self._validate_target_page(self._expected_url)

    async def _interruptible_wait(self, seconds: float) -> None:
        deadline = self._now() + seconds
        while self._now() < deadline:
            await self._checkpoint()
            await asyncio.sleep(min(0.1, max(0, deadline - self._now())))

    def _now(self) -> float:
        return time.monotonic() - self._paused_seconds

    def _stage(self, stage: str, message: str) -> None:
        self._on_stage(stage, message)

    @staticmethod
    def _data_rows(table: Locator) -> Locator:
        return table.locator('tbody tr:has(input[type="checkbox"])')

    @staticmethod
    async def _row_signature(rows: Locator) -> tuple[str, ...]:
        return tuple(text.strip() for text in await rows.all_inner_texts())

    @staticmethod
    def _normalize(value: str) -> str:
        return " ".join(value.casefold().split())

    @classmethod
    def matches_user(cls, query: str, email: str) -> bool:
        """Match an email or whole local part; a lone trailing @ is an ID delimiter."""
        query, email = cls._normalize(query), cls._normalize(email)
        if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email):
            return False
        # The observed AWS filter also accepts 'user.id@'. Treat this exactly
        # like 'user.id', never as a substring or a partial-domain email match.
        if query.endswith("@") and query.count("@") == 1:
            return bool(query[:-1]) and query[:-1] == email.split("@", 1)[0]
        return query == (email if "@" in query else email.split("@", 1)[0])

    @staticmethod
    def is_training_page(url: str) -> bool:
        parsed = urlparse(url)
        return (
            parsed.scheme == "https"
            and parsed.hostname == "skillbuilder.aws"
            and re.fullmatch(
                r"/admin/organization/modality/curriculum/training/[^/]+/?", parsed.path
            )
            is not None
            and len(parse_qs(parsed.query).get("orgId", [])) == 1
        )

    @staticmethod
    def is_same_training_destination(expected_url: str, actual_url: str) -> bool:
        """Return whether two URLs identify the same AWS training and org."""
        expected = urlparse(expected_url)
        actual = urlparse(actual_url)
        expected_org = parse_qs(expected.query).get("orgId", [])
        actual_org = parse_qs(actual.query).get("orgId", [])
        return (
            expected.scheme.casefold() == actual.scheme.casefold()
            and (expected.hostname or "").casefold()
            == (actual.hostname or "").casefold()
            and expected.path.rstrip("/") == actual.path.rstrip("/")
            and expected_org == actual_org
        )

    @staticmethod
    def _lookup_error_message(count: int) -> str:
        return (
            f"사용자 검색 결과가 {count}개입니다. 정확히 1개일 때만 할당합니다. "
            "이 사용자는 추가하지 않습니다."
        )

    @staticmethod
    async def _matched_user_label(row: Locator, query: str) -> str:
        links = [
            text.strip() for text in await row.get_by_role("link").all_inner_texts()
        ]
        for label in links:
            if re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", label):
                return label
        raise AwsAssignmentError(
            "USER_ID_UNREADABLE: 검색 결과에서 전체 이메일 주소를 읽지 못했습니다."
        )

    @staticmethod
    async def _single(locator: Locator, description: str) -> Locator:
        count = await locator.count()
        if count != 1:
            raise AwsAssignmentError(
                f"{description}이(가) 정확히 1개여야 합니다. 현재 {count}개입니다."
            )
        return locator

    @staticmethod
    async def _single_visible(locator: Locator, description: str) -> Locator:
        visible = [item for item in await locator.all() if await item.is_visible()]
        if len(visible) != 1:
            raise AwsAssignmentError(
                f"화면에 보이는 {description}이(가) 정확히 1개여야 합니다. "
                f"현재 {len(visible)}개입니다."
            )
        return visible[0]
