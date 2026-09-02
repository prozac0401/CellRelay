"""Local-page integration tests for the AWS-specific assignment workflow."""

import asyncio
import threading

import pytest
from playwright.async_api import Error as PlaywrightError
from playwright.async_api import async_playwright

from app.browser.aws_skill_builder import (
    AwsAssignmentError,
    AwsSkillBuilderAssignment,
    AwsUserLookupError,
)
from app.browser.browser_worker import BrowserWorker

_PAGE_HTML = """
<!doctype html>
<html lang="ko">
<body>
  <button id="training" aria-label="교육 할당" aria-haspopup="true">교육 할당</button>
  <div id="menu" role="menu" aria-label="교육 할당" hidden>
    <button id="to-user" role="menuitem">사용자에 할당</button>
  </div>

  <table role="grid" aria-label="사용자"><tbody id="assigned"></tbody></table>

  <div id="dialog" role="dialog" aria-label="사용자 선택" hidden>
    <input id="search" placeholder="사용자 찾기">
    <table role="table" aria-label="사용자"><tbody id="results"></tbody></table>
    <button id="assign" data-testid="assign-classroom-training-assign-modal-btn"
            disabled>할당</button>
  </div>

  <div id="confirmation" role="dialog" aria-label="교육 할당 확인"
       data-testid="assign_classroom_training_confirmation_modal" hidden>
    <p>선택한 사용자를 교육에 할당하시겠습니까?</p>
    <button id="confirm">확인</button>
  </div>

  <script>
    const users = [
      "target@example.com",
      "second@example.com",
      "outside@sample.net",
    ];
    const menu = document.querySelector("#menu");
    const dialog = document.querySelector("#dialog");
    const search = document.querySelector("#search");
    const results = document.querySelector("#results");
    const assign = document.querySelector("#assign");
    const confirmation = document.querySelector("#confirmation");
    const confirm = document.querySelector("#confirm");

    function renderRows(values) {
      results.replaceChildren();
      for (const value of values) {
        const row = document.createElement("tr");
        row.innerHTML = `<td><input type="checkbox" aria-label="${value}"></td>`
          + `<td><a href="#${value}">${value}</a></td>`;
        row.querySelector("input").addEventListener("change", event => {
          assign.disabled = !event.target.checked;
        });
        results.append(row);
      }
    }

    renderRows(users);
    document.querySelector("#training").addEventListener("click", () => {
      menu.hidden = false;
    });
    document.querySelector("#to-user").addEventListener("click", () => {
      menu.hidden = true;
      dialog.hidden = false;
    });
    search.addEventListener("keydown", event => {
      if (event.key !== "Enter") return;
      const query = search.value.toLowerCase();
      setTimeout(() => renderRows(users.filter(user => user.includes(query))), 80);
    });
    assign.addEventListener("click", () => {
      const selected = results.querySelector("input:checked")
        .closest("tr").querySelector("a").textContent;
      dialog.hidden = true;
      confirmation.hidden = false;
      const row = document.createElement("tr");
      row.innerHTML = `<td><a>${selected}</a></td>`;
      document.querySelector("#assigned").append(row);
    });
    confirm.addEventListener("click", () => {
      confirmation.hidden = true;
    });
  </script>
</body>
</html>
"""

_PORTALED_DELAYED_SEARCH_HTML = (
    _PAGE_HTML.replace(
        '    <input id="search" placeholder="사용자 찾기">\n',
        "",
    )
    .replace(
        '  <div id="dialog" role="dialog" aria-label="사용자 선택" hidden>\n',
        '  <input id="search" placeholder="사용자 찾기" hidden>\n\n'
        '  <div id="dialog" role="dialog" aria-label="사용자 선택" hidden>\n',
    )
    .replace(
        "      dialog.hidden = false;\n",
        "      dialog.hidden = false;\n"
        "      setTimeout(() => { search.hidden = false; }, 300);\n",
    )
)


async def _launch_edge(playwright):
    try:
        return await playwright.chromium.launch(channel="msedge", headless=True)
    except PlaywrightError as exc:
        pytest.skip(f"Microsoft Edge is unavailable: {exc}")


def test_assigns_only_single_matching_user_and_verifies_main_grid() -> None:
    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await _launch_edge(playwright)
            try:
                page = await browser.new_page()
                await page.set_content(_PAGE_HTML)
                stages: list[str] = []
                workflow = AwsSkillBuilderAssignment(
                    page,
                    threading.Event(),
                    threading.Event(),
                    on_stage=lambda stage, _message: stages.append(stage),
                    lookup_timeout_ms=2_000,
                    result_stable_ms=300,
                    invalid_result_stable_ms=500,
                    verification_timeout_ms=2_000,
                    enforce_aws_host=False,
                )

                result = await workflow.assign_user("target@example.com")

                assert result.matched_user_label == "target@example.com"
                assert await page.get_by_role(
                    "dialog",
                    name="사용자 선택",
                ).is_hidden()
                assert await page.get_by_test_id(
                    "assign_classroom_training_confirmation_modal"
                ).is_hidden()
                assert (
                    await page.get_by_role("grid", name="사용자")
                    .get_by_text("target@example.com", exact=True)
                    .count()
                    == 1
                )
                assert stages == [
                    "OPENING_ASSIGNMENT",
                    "SEARCHING_USER",
                    "SELECTING_USER",
                    "ASSIGNING_USER",
                    "CONFIRMING_ASSIGNMENT",
                    "VERIFYING_ASSIGNMENT",
                ]
            finally:
                await browser.close()

    asyncio.run(run())


def test_confirmation_modal_is_closed_before_the_next_assignment() -> None:
    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await _launch_edge(playwright)
            try:
                page = await browser.new_page()
                await page.set_content(_PAGE_HTML)
                workflow = AwsSkillBuilderAssignment(
                    page,
                    threading.Event(),
                    threading.Event(),
                    lookup_timeout_ms=2_000,
                    result_stable_ms=300,
                    invalid_result_stable_ms=500,
                    verification_timeout_ms=2_000,
                    enforce_aws_host=False,
                )

                await workflow.assign_user("target@example.com")
                await workflow.assign_user("second@example.com")

                assert await page.get_by_test_id(
                    "assign_classroom_training_confirmation_modal"
                ).is_hidden()
                assert (
                    await page.get_by_role("grid", name="사용자")
                    .locator("tbody tr")
                    .count()
                    == 2
                )
            finally:
                await browser.close()

    asyncio.run(run())


def test_finds_delayed_search_input_rendered_outside_dialog() -> None:
    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await _launch_edge(playwright)
            try:
                page = await browser.new_page()
                await page.set_content(_PORTALED_DELAYED_SEARCH_HTML)
                workflow = AwsSkillBuilderAssignment(
                    page,
                    threading.Event(),
                    threading.Event(),
                    lookup_timeout_ms=2_000,
                    result_stable_ms=300,
                    invalid_result_stable_ms=500,
                    readiness_timeout_ms=2_000,
                    verification_timeout_ms=2_000,
                    enforce_aws_host=False,
                )

                result = await workflow.assign_user("target@example.com")

                assert result.matched_user_label == "target@example.com"
                assert await page.get_by_test_id(
                    "assign_classroom_training_confirmation_modal"
                ).is_hidden()
            finally:
                await browser.close()

    asyncio.run(run())


def test_readiness_check_finds_training_button_without_clicking() -> None:
    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await _launch_edge(playwright)
            try:
                page = await browser.new_page()
                await page.set_content(_PAGE_HTML)
                workflow = AwsSkillBuilderAssignment(
                    page,
                    threading.Event(),
                    threading.Event(),
                    enforce_aws_host=False,
                )

                assert await workflow.validate_ready() == page.url
                assert await page.get_by_role(
                    "menu",
                    name="교육 할당",
                ).is_hidden()
                assert await page.get_by_role(
                    "dialog",
                    name="사용자 선택",
                ).is_hidden()
            finally:
                await browser.close()

    asyncio.run(run())


def test_readiness_check_waits_for_training_button_to_render() -> None:
    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await _launch_edge(playwright)
            try:
                page = await browser.new_page()
                await page.set_content(
                    """
                    <main>교육 상세 페이지를 불러오는 중</main>
                    <script>
                      setTimeout(() => {
                        const button = document.createElement("button");
                        button.setAttribute("aria-label", "교육 할당");
                        button.textContent = "교육 할당";
                        document.body.append(button);
                      }, 300);
                    </script>
                    """
                )
                workflow = AwsSkillBuilderAssignment(
                    page,
                    threading.Event(),
                    threading.Event(),
                    readiness_timeout_ms=2_000,
                    enforce_aws_host=False,
                )

                assert await workflow.validate_ready() == page.url
            finally:
                await browser.close()

    asyncio.run(run())


def test_readiness_check_rejects_page_without_training_button() -> None:
    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await _launch_edge(playwright)
            try:
                page = await browser.new_page()
                await page.set_content("<main>로그인 또는 로딩 화면</main>")
                workflow = AwsSkillBuilderAssignment(
                    page,
                    threading.Event(),
                    threading.Event(),
                    readiness_timeout_ms=1_000,
                    enforce_aws_host=False,
                )

                with pytest.raises(AwsAssignmentError, match="교육 할당 버튼"):
                    await workflow.validate_ready()
            finally:
                await browser.close()

    asyncio.run(run())


def test_worker_follows_login_created_tab_for_expected_training() -> None:
    expected_url = (
        "https://skillbuilder.aws/admin/organization/modality/curriculum/"
        "training/5d73636f-9539-4532-92df-94511a3c4bda"
        "?orgId=9377416d-12ef-431d-a739-7a754f3321ba"
    )

    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await _launch_edge(playwright)
            try:
                context = await browser.new_context()
                await context.route(
                    "https://skillbuilder.aws/**",
                    lambda route: route.fulfill(body=_PAGE_HTML),
                )
                login_page = await context.new_page()
                await login_page.set_content("<main>로그인 화면</main>")
                training_page = await context.new_page()
                await training_page.goto(expected_url)

                worker = BrowserWorker()
                worker._context = context
                worker._page = login_page

                assert await worker._select_target_page(expected_url) is training_page
                assert worker._page is training_page
            finally:
                await browser.close()

    asyncio.run(run())


def test_worker_waits_for_login_created_training_tab() -> None:
    expected_url = (
        "https://skillbuilder.aws/admin/organization/modality/curriculum/"
        "training/5d73636f-9539-4532-92df-94511a3c4bda"
        "?orgId=9377416d-12ef-431d-a739-7a754f3321ba"
    )

    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await _launch_edge(playwright)
            try:
                context = await browser.new_context()
                await context.route(
                    "https://skillbuilder.aws/**",
                    lambda route: route.fulfill(body=_PAGE_HTML),
                )
                login_page = await context.new_page()
                await login_page.set_content("<main>로그인 화면</main>")
                worker = BrowserWorker()
                worker._context = context
                worker._page = login_page

                async def open_training_after_login():
                    await asyncio.sleep(0.3)
                    training_page = await context.new_page()
                    await training_page.goto(expected_url)
                    return training_page

                pending_page = asyncio.create_task(open_training_after_login())
                selected_page = await worker._select_target_page(
                    expected_url,
                    timeout_ms=2_000,
                )

                assert selected_page is await pending_page
                assert worker._page is selected_page
            finally:
                await browser.close()

    asyncio.run(run())


def test_refuses_to_assign_when_two_users_match() -> None:
    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await _launch_edge(playwright)
            try:
                page = await browser.new_page()
                await page.set_content(_PAGE_HTML)
                workflow = AwsSkillBuilderAssignment(
                    page,
                    threading.Event(),
                    threading.Event(),
                    lookup_timeout_ms=2_000,
                    result_stable_ms=300,
                    invalid_result_stable_ms=500,
                    verification_timeout_ms=2_000,
                    enforce_aws_host=False,
                )

                with pytest.raises(AwsUserLookupError, match="2개"):
                    await workflow.assign_user("@example.com")

                assert await page.get_by_test_id(
                    "assign-classroom-training-assign-modal-btn"
                ).is_disabled()
                assert (
                    await page.get_by_role("grid", name="사용자")
                    .locator("tbody tr")
                    .count()
                    == 0
                )
            finally:
                await browser.close()

    asyncio.run(run())
