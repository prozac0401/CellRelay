"""State-machine primitives shared by the controller and UI."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class AppState(str, Enum):
    """All user-visible CellRelay states."""

    IDLE = "IDLE"
    LOADING_EXCEL = "LOADING_EXCEL"
    READY = "READY"
    INPUTTING = "INPUTTING"
    WAITING_FOR_CLEAR = "WAITING_FOR_CLEAR"
    PAUSED = "PAUSED"
    STOPPING = "STOPPING"
    COMPLETED = "COMPLETED"
    ERROR = "ERROR"


class StateTransitionError(RuntimeError):
    """Raised when code attempts an invalid state transition."""


_ALLOWED_TRANSITIONS: dict[AppState, set[AppState]] = {
    AppState.IDLE: {AppState.LOADING_EXCEL, AppState.ERROR},
    AppState.LOADING_EXCEL: {AppState.IDLE, AppState.READY, AppState.ERROR},
    AppState.READY: {
        AppState.IDLE,
        AppState.LOADING_EXCEL,
        AppState.INPUTTING,
        AppState.COMPLETED,
        AppState.ERROR,
    },
    AppState.INPUTTING: {
        AppState.STOPPING,
        AppState.WAITING_FOR_CLEAR,
        AppState.PAUSED,
        AppState.READY,
        AppState.COMPLETED,
        AppState.ERROR,
    },
    AppState.WAITING_FOR_CLEAR: {
        AppState.STOPPING,
        AppState.INPUTTING,
        AppState.PAUSED,
        AppState.READY,
        AppState.COMPLETED,
        AppState.ERROR,
    },
    AppState.PAUSED: {
        AppState.STOPPING,
        AppState.INPUTTING,
        AppState.WAITING_FOR_CLEAR,
        AppState.READY,
        AppState.COMPLETED,
        AppState.ERROR,
    },
    AppState.COMPLETED: {
        AppState.IDLE,
        AppState.LOADING_EXCEL,
        AppState.READY,
        AppState.INPUTTING,
        AppState.ERROR,
    },
    AppState.STOPPING: {AppState.READY, AppState.ERROR},
    AppState.ERROR: {
        AppState.IDLE,
        AppState.LOADING_EXCEL,
        AppState.READY,
    },
}


class StateMachine:
    """Small explicit state machine without UI dependencies."""

    def __init__(self) -> None:
        self._state = AppState.IDLE

    @property
    def state(self) -> AppState:
        return self._state

    def transition_to(self, new_state: AppState) -> bool:
        """Move to *new_state* and return whether the state changed."""
        if new_state == self._state:
            return False
        if new_state not in _ALLOWED_TRANSITIONS[self._state]:
            raise StateTransitionError(
                f"Invalid transition: {self._state.value} -> {new_state.value}"
            )
        self._state = new_state
        return True


@dataclass(slots=True)
class ProgressSnapshot:
    """The compact progress model rendered by the main window."""

    current_cell: str = "-"
    current_value: str = ""
    processed_count: int = 0
    skipped_count: int = 0
    total_items: int = 0
    last_message: str = "대기 중입니다."
