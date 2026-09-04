"""Observe the request caused by Enter without guessing private AWS API routes."""

from urllib.parse import unquote

from playwright.async_api import Page, Request


class SearchObserver:
    def __init__(self, page: Page, query: str) -> None:
        self.page = page
        self.query = query.casefold()
        self.pending: set[Request] = set()
        self.completed = False
        self.failed = False

    def start(self) -> None:
        self.page.on("request", self._request)
        self.page.on("requestfinished", self._finished)
        self.page.on("requestfailed", self._failed)
        self.page.on("response", self._response)

    def close(self) -> None:
        for name, handler in (
            ("request", self._request),
            ("requestfinished", self._finished),
            ("requestfailed", self._failed),
            ("response", self._response),
        ):
            self.page.remove_listener(name, handler)

    def _request(self, request: Request) -> None:
        if request.resource_type not in {"xhr", "fetch"}:
            return
        payload = unquote(unquote(request.url + " " + (request.post_data or "")))
        if self.query in payload.casefold():
            self.pending.add(request)

    def _response(self, response) -> None:
        if response.request in self.pending and not 200 <= response.status < 300:
            self.failed = True

    def _finished(self, request: Request) -> None:
        if request in self.pending:
            self.pending.remove(request)
            self.completed = True

    def _failed(self, request: Request) -> None:
        if request in self.pending:
            self.pending.remove(request)
            self.failed = True
