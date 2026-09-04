"""Bilingual integration regressions on local fixtures, never live assignments."""

import asyncio
import threading

import pytest
from playwright.async_api import async_playwright
from test_aws_skill_builder import _PAGE_HTML

from app.browser.aws_skill_builder import AwsAssignmentError, AwsSkillBuilderAssignment
from app.browser.browser_worker import BrowserWorker


def localized_html(locale):
    html = _PAGE_HTML
    if locale == "en":
        for ko, en in [
            (
                "선택한 모든 사용자를 등록하고 싶습니다.",
                "I want to register all selected users.",
            ),
            ("사용자를 등록하시겠습니까?", "Do you want to register users?"),
            ("사용자에 할당", "Assign to users"),
            ("사용자 선택", "Select users"),
            ("사용자 찾기", "Find users"),
            ("교육 할당 확인", "Assign training confirmation"),
            ("교육 할당", "Assign training"),
            ("일치 항목 없음", "No matches"),
            ("등록 상태", "Registration status"),
            ("등록 대기", "Pending"),
            ("등록됨", "Registered"),
            ("사용자", "Users"),
            ("취소", "Cancel"),
            ("할당", "Assign"),
            ("완료", "Done"),
        ]:
            html = html.replace(ko, en)
    return html


def workflow(page, **kwargs):
    options = {
        "lookup_timeout_ms": 2000,
        "result_stable_ms": 300,
        "readiness_timeout_ms": 1000,
        "verification_timeout_ms": 1000,
        "post_confirmation_settle_ms": 0,
        "post_cancel_settle_ms": 0,
        "enforce_aws_host": False,
    }
    options.update(kwargs)
    return AwsSkillBuilderAssignment(
        page, threading.Event(), threading.Event(), **options
    )


@pytest.mark.parametrize("locale", ["ko", "en"])
def test_bilingual_assignment_skip_and_next_button_reacquisition(locale):
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                page = await browser.new_page()
                await page.set_content(localized_html(locale))
                w = workflow(page)
                await w.validate_ready()
                assert (await w.assign_user("target")).assigned
                assert not (await w.assign_user("missing")).assigned
                assert (await w.assign_user("second@example.com")).assigned
                mismatch = await w.assign_user("side")
                assert (
                    not mismatch.assigned
                    and "IDENTITY_MISMATCH" in mismatch.skip_reason
                )
                multiple = await w.assign_user("@example.com")
                assert (
                    not multiple.assigned and "MULTIPLE_MATCHES" in multiple.skip_reason
                )
                assert await page.locator("#assigned tr").count() == 2
            finally:
                await browser.close()

    asyncio.run(run())


@pytest.mark.parametrize("locale", ["ko", "en"])
def test_no_success_for_pending_enrollment_or_next_button_timeout(locale):
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                page = await browser.new_page()
                # Registration does not complete: an email row alone is not success.
                html = localized_html(locale)
                status = "등록됨" if locale == "ko" else "Registered"
                await page.set_content(
                    html.replace(
                        f".textContent = '{status}'", ".textContent = 'Pending'"
                    )
                )
                with pytest.raises(AwsAssignmentError, match="ENROLLMENT_UNCONFIRMED"):
                    await workflow(page).assign_user("target")
                page = await browser.new_page()
                # Successful registration must return even when the *next*
                # training button never comes back. Fail only the next row.
                await page.set_content(
                    html.replace(
                        "document.body.prepend(trainingButton);",
                        "/* removed permanently */",
                    )
                )
                w = workflow(page)
                assert (await w.assign_user("target")).assigned
                with pytest.raises(AwsAssignmentError, match="교육 할당 버튼"):
                    await w.assign_user("second")
                assert await page.locator("#assigned tr").count() == 1
            finally:
                await browser.close()

    asyncio.run(run())


def network_html():
    return _PAGE_HTML.replace(
        "setTimeout(() => {\n        renderRows",
        "fetch('https://search.fixture/users?q=' + encodeURIComponent(query)).then(() => {\n        renderRows",
    ).replace("      }, 300);", "      });")


def test_slow_failed_and_unconfirmed_searches_are_not_skipped():
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                for delay, status, query, expected in [
                    (3.3, 200, "target", "success"),
                    (0.1, 500, "missing", "SEARCH_FAILED"),
                    (2, 200, "missing", "SEARCH_TIMEOUT"),
                ]:
                    page = await browser.new_page()

                    async def respond(route, _request, *, delay=delay, status=status):
                        await asyncio.sleep(delay)
                        await route.fulfill(
                            status=status,
                            body="{}",
                            headers={"Access-Control-Allow-Origin": "*"},
                        )

                    await page.route("https://search.fixture/**", respond)
                    await page.set_content(network_html())
                    w = workflow(
                        page, lookup_timeout_ms=5000 if expected == "success" else 1000
                    )
                    if expected == "success":
                        assert (await w.assign_user(query)).assigned
                    else:
                        with pytest.raises(AwsAssignmentError, match=expected):
                            await w.assign_user(query)
                        assert await page.locator("#assigned tr").count() == 0
                    await page.close()
            finally:
                await browser.close()

    asyncio.run(run())


def test_pause_does_not_consume_search_timeout():
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                page = await browser.new_page()
                await page.route(
                    "https://search.fixture/**",
                    lambda route: route.fulfill(
                        body="{}", headers={"Access-Control-Allow-Origin": "*"}
                    ),
                )
                await page.set_content(network_html())
                w = workflow(page, lookup_timeout_ms=1000)
                original_wait = w._wait_for_one_matching_row

                async def pause_then_wait(*args):
                    w._pause_event.set()

                    async def resume():
                        await asyncio.sleep(1.4)
                        w._pause_event.clear()

                    task = asyncio.create_task(resume())
                    try:
                        return await original_wait(*args)
                    finally:
                        await task

                w._wait_for_one_matching_row = pause_then_wait
                assert (await w.assign_user("target")).assigned
            finally:
                await browser.close()

    asyncio.run(run())


def test_login_discovery_rejects_multiple_training_tabs():
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                context = await browser.new_context()
                await context.route(
                    "https://skillbuilder.aws/**",
                    lambda route: route.fulfill(body=_PAGE_HTML),
                )
                login = await context.new_page()
                await login.goto("https://skillbuilder.aws/login")
                training = await context.new_page()
                await training.goto(
                    "https://skillbuilder.aws/admin/organization/modality/curriculum/training/one?orgId=org"
                )
                worker = BrowserWorker()
                worker._context, worker._page = context, login
                assert await worker._select_target_page("") is training
                other = await context.new_page()
                await other.goto(
                    "https://skillbuilder.aws/admin/organization/modality/curriculum/training/two?orgId=org"
                )
                with pytest.raises(RuntimeError, match="여러 개"):
                    await worker._select_target_page("")
            finally:
                await browser.close()

    asyncio.run(run())


@pytest.mark.parametrize(
    "query,email,expected",
    [
        ("kang", "jeong.kang@example.com", False),
        ("jeong.kang", "jeong.kang2@example.com", False),
        (" jeong.kang ", "JEONG.KANG@example.com", True),
        ("a@one.com", "a@two.com", False),
        ("a@one.com", "a@one.com", True),
        ("Example", "a@example.com", False),
    ],
)
def test_identity_matching(query, email, expected):
    assert AwsSkillBuilderAssignment.matches_user(query, email) == expected
