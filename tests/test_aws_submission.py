"""Submit and continue without consulting the training-detail roster."""

import asyncio
import re
import time

import pytest
from playwright.async_api import async_playwright
from test_aws_safety import localized_html, workflow

from app.browser.aws_skill_builder import AwsAssignmentCancelled, AwsAssignmentError
from app.browser.browser_worker import BrowserWorker


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


async def set_observed_submission_page(page, html):
    """Keep control tests independent of catching a brief loading indicator."""
    await page.route(
        "https://search.fixture/**",
        lambda route: route.fulfill(
            body="{}", headers={"Access-Control-Allow-Origin": "*"}
        ),
    )
    html = html.replace(
        "setTimeout(() => {\n        renderRows",
        "fetch('https://search.fixture/users?q=' + encodeURIComponent(query))"
        ".then(() => {\n        renderRows",
    ).replace("      }, 300);", "      });")
    await page.set_content(html)


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


@pytest.mark.parametrize("stop_point", ["before_done", "before_closed"])
def test_stop_before_submission_evidence_does_not_report_completion(stop_point):
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                page = await browser.new_page()
                html = submission_html("en", "absent")
                if stop_point == "before_closed":
                    html = html.replace(
                        "confirmation.hidden = true;", "/* popup stays open */"
                    )
                await set_observed_submission_page(page, html)
                w = workflow(
                    page,
                    lookup_timeout_ms=15_000,
                    readiness_timeout_ms=20_000,
                    post_confirmation_settle_ms=300,
                )
                original_enabled = w._wait_until_enabled
                original_hidden = w._wait_until_hidden
                stages = []
                w._on_stage = lambda stage, _message: stages.append(stage)

                async def stop_before_done(locator, timeout_ms):
                    await original_enabled(locator, timeout_ms)
                    if await locator.get_attribute("id") == "confirm":
                        w._stop_event.set()

                async def stop_before_closed(locator, timeout_ms, *args):
                    if args == ("최종 할당 후 확인 창이 닫히지 않았습니다.",):
                        assert await page.locator("#confirmation").is_visible()
                        assert await page.evaluate("window.doneChecks") == [True]
                        w._stop_event.set()
                    await original_hidden(locator, timeout_ms, *args)

                if stop_point == "before_done":
                    w._wait_until_enabled = stop_before_done
                else:
                    w._wait_until_hidden = stop_before_closed

                with pytest.raises(AwsAssignmentCancelled):
                    await w.assign_user("target")
                assert "ASSIGNMENT_SUBMITTED" not in stages
                expected_clicks = [] if stop_point == "before_done" else [True]
                assert await page.evaluate("window.doneChecks") == expected_clicks
                assert await page.locator("#confirmation").is_visible()
            finally:
                await browser.close()

    asyncio.run(run())


@pytest.mark.parametrize("control", ["stop", "pause_then_stop", "pause"])
def test_confirmed_submission_reaches_worker_after_control_during_settle(
    control, monkeypatch
):
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                page = await browser.new_page()
                await set_observed_submission_page(
                    page, submission_html("en", "absent")
                )
                worker = BrowserWorker()
                events = []
                handles = []
                closed_at = None
                worker.assignment_completed.connect(
                    lambda run_id: events.append(("completed", run_id))
                )
                worker.operation_cancelled.connect(
                    lambda run_id, _operation: events.append(("cancelled", run_id))
                )

                async def select_page(_expected_url):
                    return page

                def make_workflow(page, stop_event, pause_event, on_stage):
                    w = workflow(
                        page,
                        lookup_timeout_ms=15_000,
                        readiness_timeout_ms=20_000,
                        post_confirmation_settle_ms=300,
                    )
                    w._stop_event, w._pause_event = stop_event, pause_event
                    w._on_stage = on_stage
                    original_hidden = w._wait_until_hidden

                    async def schedule_control_after_closed(locator, timeout_ms, *args):
                        nonlocal closed_at
                        await original_hidden(locator, timeout_ms, *args)
                        if args == ("최종 할당 후 확인 창이 닫히지 않았습니다.",):
                            assert await page.locator("#confirmation").is_hidden()
                            closed_at = time.monotonic()
                            loop = asyncio.get_running_loop()
                            if control.startswith("pause"):
                                handles.append(loop.call_later(0.05, worker.request_pause))
                            if control in {"stop", "pause_then_stop"}:
                                handles.append(loop.call_later(0.1, worker.request_stop))

                    w._wait_until_hidden = schedule_control_after_closed
                    return w

                worker._select_target_page = select_page
                monkeypatch.setattr(
                    "app.browser.browser_worker.AwsSkillBuilderAssignment", make_workflow
                )
                try:
                    await asyncio.wait_for(worker._assign_aws_user(7, "target", ""), 60)
                    assert closed_at is not None
                    elapsed = time.monotonic() - closed_at
                    assert events == [("completed", 7)]
                    assert await page.evaluate("window.doneChecks") == [True]
                    assert await page.evaluate("window.submittedSelections") == [
                        "target@example.com"
                    ]
                    assert await page.locator("#confirmation").is_hidden()
                    if control == "pause":
                        assert worker._pause_event.is_set()
                        assert elapsed >= 0.28
                    else:
                        assert worker._stop_event.is_set()
                finally:
                    for handle in handles:
                        handle.cancel()
            finally:
                await browser.close()

    asyncio.run(run())
