"""Local-page integration tests for the AWS-specific assignment workflow."""

import threading

import pytest
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from app.browser.aws_skill_builder import (
    AwsAssignmentError,
    AwsSkillBuilderAssignment,
    AwsUserLookupError,
)

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
    <input id="search" role="combobox" aria-label="사용자 찾기">
    <table role="table" aria-label="사용자"><tbody id="results"></tbody></table>
    <button id="assign" data-testid="assign-classroom-training-assign-modal-btn"
            disabled>할당</button>
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
    search.addEventListener("input", () => {
      const query = search.value.toLowerCase();
      setTimeout(() => renderRows(users.filter(user => user.includes(query))), 80);
    });
    assign.addEventListener("click", () => {
      const selected = results.querySelector("input:checked")
        .closest("tr").querySelector("a").textContent;
      dialog.hidden = true;
      const row = document.createElement("tr");
      row.innerHTML = `<td><a>${selected}</a></td>`;
      document.querySelector("#assigned").append(row);
    });
  </script>
</body>
</html>
"""


def _launch_edge(playwright):
    try:
        return playwright.chromium.launch(channel="msedge", headless=True)
    except PlaywrightError as exc:
        pytest.skip(f"Microsoft Edge is unavailable: {exc}")


def test_assigns_only_single_matching_user_and_verifies_main_grid() -> None:
    with sync_playwright() as playwright:
        browser = _launch_edge(playwright)
        try:
            page = browser.new_page()
            page.set_content(_PAGE_HTML)
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

            result = workflow.assign_user("target@example.com")

            assert result.matched_user_label == "target@example.com"
            assert page.get_by_role("dialog", name="사용자 선택").is_hidden()
            assert (
                page.get_by_role("grid", name="사용자")
                .get_by_text("target@example.com", exact=True)
                .count()
                == 1
            )
            assert stages == [
                "OPENING_ASSIGNMENT",
                "SEARCHING_USER",
                "SELECTING_USER",
                "ASSIGNING_USER",
                "VERIFYING_ASSIGNMENT",
            ]
        finally:
            browser.close()


def test_readiness_check_finds_training_button_without_clicking() -> None:
    with sync_playwright() as playwright:
        browser = _launch_edge(playwright)
        try:
            page = browser.new_page()
            page.set_content(_PAGE_HTML)
            workflow = AwsSkillBuilderAssignment(
                page,
                threading.Event(),
                threading.Event(),
                enforce_aws_host=False,
            )

            assert workflow.validate_ready() == page.url
            assert page.get_by_role("menu", name="교육 할당").is_hidden()
            assert page.get_by_role("dialog", name="사용자 선택").is_hidden()
        finally:
            browser.close()


def test_readiness_check_rejects_page_without_training_button() -> None:
    with sync_playwright() as playwright:
        browser = _launch_edge(playwright)
        try:
            page = browser.new_page()
            page.set_content("<main>로그인 또는 로딩 화면</main>")
            workflow = AwsSkillBuilderAssignment(
                page,
                threading.Event(),
                threading.Event(),
                enforce_aws_host=False,
            )

            with pytest.raises(AwsAssignmentError, match="교육 할당 버튼"):
                workflow.validate_ready()
        finally:
            browser.close()


def test_refuses_to_assign_when_two_users_match() -> None:
    with sync_playwright() as playwright:
        browser = _launch_edge(playwright)
        try:
            page = browser.new_page()
            page.set_content(_PAGE_HTML)
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
                workflow.assign_user("@example.com")

            assert page.get_by_test_id(
                "assign-classroom-training-assign-modal-btn"
            ).is_disabled()
            assert (
                page.get_by_role("grid", name="사용자").locator("tbody tr").count() == 0
            )
        finally:
            browser.close()
