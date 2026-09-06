"""Submit and continue without consulting the training-detail roster."""

import asyncio
import re

import pytest
from playwright.async_api import async_playwright
from test_aws_safety import localized_html, workflow

from app.browser.aws_skill_builder import AwsAssignmentError


def submission_html(locale, roster):
    html = localized_html(locale)
    status = "등록됨" if locale == "ko" else "Proxy-enrolled"
    # The UI can submit users without adding them to the displayed roster page.
    # Track synthetic submissions separately from any main-table DOM.
    html = (
        html.replace(
            "    const users = [",
            "    window.submittedSelections = []; window.doneChecks = [];\n"
            "    const users = [",
        )
        .replace(
            'document.querySelector("#assigned").append(row);',
            "window.submittedSelections.push(selected);",
        )
        .replace(
            f"document.querySelector('#assigned tr:last-child td:last-child').textContent = '{status}';",
            "window.doneChecks.push(registerAll.checked);",
        )
    )
    # Only the search popup retains its table. The detail roster is irrelevant.
    replacement = ""
    if roster != "absent":
        name = "사용자" if locale == "ko" else "Users"
        header = "등록 상태" if locale == "ko" else "Enrollment status"
        email = "other@example.com" if roster == "other_page" else "target@example.com"
        state = status if roster == "other_page" else "Pending"
        replacement = (
            f'<table role="grid" aria-label="{name}"><thead><tr>'
            f"<th>Email</th><th>{header}</th></tr></thead><tbody>"
            f'<tr><td><a href="#">{email}</a></td><td>{state}</td></tr>'
            '</tbody></table><button aria-label="Next page">2</button>'
        )
    return re.sub(r'  <table role="grid"[^\n]+</table>', replacement, html, count=1)


@pytest.mark.parametrize("locale", ["ko", "en"])
@pytest.mark.parametrize("roster", ["absent", "other_page", "pending"])
def test_two_submissions_ignore_missing_or_stale_detail_roster(locale, roster):
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                page = await browser.new_page()
                await page.set_content(submission_html(locale, roster))
                w = workflow(page, post_confirmation_settle_ms=50)
                first = await w.assign_user("target@")
                second = await w.assign_user("second@example.com")
                assert first.assigned and second.assigned
                assert await page.evaluate("window.submittedSelections") == [
                    "target@example.com",
                    "second@example.com",
                ]
                assert await page.evaluate("window.doneChecks") == [True, True]
                assert await page.locator("#confirmation").is_hidden()
            finally:
                await browser.close()

    asyncio.run(run())


@pytest.mark.parametrize("locale", ["ko", "en"])
def test_confirmation_must_close_before_submission_completes(locale):
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                page = await browser.new_page()
                html = submission_html(locale, "absent").replace(
                    "confirmation.hidden = true;", "/* popup stays open */"
                )
                await page.set_content(html)
                w = workflow(page)
                original = w._wait_until_hidden

                async def short_wait(locator, timeout_ms, *args):
                    return await original(locator, min(timeout_ms, 500), *args)

                w._wait_until_hidden = short_wait
                with pytest.raises(AwsAssignmentError, match="확인 창이 닫히지"):
                    await w.assign_user("target")
                assert await page.evaluate("window.doneChecks") == [True]
            finally:
                await browser.close()

    asyncio.run(run())


@pytest.mark.parametrize("locale", ["ko", "en"])
def test_waits_for_confirmation_close_and_settling_delay(locale):
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                page = await browser.new_page()
                html = submission_html(locale, "absent").replace(
                    "confirmation.hidden = true;",
                    "setTimeout(() => { confirmation.hidden = true; "
                    "window.closedAt = performance.now(); }, 250);",
                )
                await page.set_content(html)
                w = workflow(page, post_confirmation_settle_ms=300)
                assert (await w.assign_user("target")).assigned
                assert await page.locator("#confirmation").is_hidden()
                assert await page.evaluate("performance.now() - window.closedAt") >= 280
                assert await page.evaluate("window.doneChecks") == [True]
            finally:
                await browser.close()

    asyncio.run(run())
