"""Screenshot-derived English labels with synthetic users and mocked DOM structure."""

import asyncio

import pytest
from playwright.async_api import async_playwright
from test_aws_safety import localized_html, workflow


def observed_english_html(user_status: str, enrollment_status: str) -> str:
    html = localized_html("en")
    # The screenshot separates account status from enrollment status, and the
    # latter is not the last column. Do not infer DOM/ARIA from pixels: retain
    # the existing known roles and exercise the observed visible column order.
    html = (
        html.replace(
            "<th>이메일</th><th>Enrollment status</th>",
            "<th><input type='checkbox' aria-label='Select all users'></th>"
            "<th>Email</th><th>Name</th><th>User status</th><th>Enrollment status</th>"
            "<th>Assigned on</th><th>Enrolled on</th><th>Withdrawn on</th>",
        )
        .replace(
            '<td><a href="#${selected}">${selected}</a></td><td>Pending</td>',
            '<td><input type="checkbox"></td>'
            '<td><a class="email" href="#${selected}">${selected}</a></td>'
            f'<td>Example Learner</td><td>{user_status}</td><td class="enrollment">Pending</td>'
            "<td>-</td><td>September 5, 2026</td><td>-</td>",
        )
        .replace(
            "#assigned tr:last-child td:last-child",
            "#assigned tr:last-child .enrollment",
        )
        .replace(
            ".textContent = 'Proxy-enrolled'", f".textContent = '{enrollment_status}'"
        )
    )
    # In the observed confirmation, Done is already enabled before checking.
    html = html.replace('id="confirm" disabled', 'id="confirm"').replace(
        "confirm.disabled = true;", "confirm.disabled = false;"
    )
    return html.replace(
        "</body>",
        """
      <style>.email { display:block; width:75px; overflow:hidden;
                      white-space:nowrap; text-overflow:ellipsis; }</style>
      <script>
        window.enrollmentCheckedOnDone = false;
        document.querySelector('#confirm').addEventListener('click', () => {
          window.enrollmentCheckedOnDone = document.querySelector('#register-all').checked;
        });
      </script>
    </body>""",
    )


@pytest.mark.parametrize(
    "user_status,enrollment_status",
    [("Invited", "Proxy-enrolled"), ("Active", "Pending")],
)
def test_enabled_done_requires_checkbox_but_roster_status_is_ignored(
    user_status, enrollment_status
):
    async def run():
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=True)
            try:
                page = await browser.new_page()
                await page.set_content(
                    observed_english_html(user_status, enrollment_status)
                )
                w = workflow(page)
                result = await w.assign_user("target@")
                assert result.assigned
                assert result.matched_user_label == "target@example.com"
                assert await page.evaluate("window.enrollmentCheckedOnDone")
                assert (
                    await page.locator("#assigned .enrollment").inner_text()
                    == enrollment_status
                )
                assert await page.locator("#assigned th").count() == 0
                assert await page.get_by_role("columnheader").count() == 8
            finally:
                await browser.close()

    asyncio.run(run())
