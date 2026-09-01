import pytest

from app.core.state import AppState, StateMachine, StateTransitionError


def test_normal_workflow_transitions() -> None:
    state = StateMachine()
    for next_state in (
        AppState.LOADING_EXCEL,
        AppState.READY,
        AppState.INPUTTING,
        AppState.WAITING_FOR_CLEAR,
        AppState.INPUTTING,
        AppState.WAITING_FOR_CLEAR,
        AppState.COMPLETED,
    ):
        assert state.transition_to(next_state)
    assert state.state is AppState.COMPLETED


def test_pause_and_resume_preserve_explicit_phase() -> None:
    state = StateMachine()
    state.transition_to(AppState.LOADING_EXCEL)
    state.transition_to(AppState.READY)
    state.transition_to(AppState.INPUTTING)
    state.transition_to(AppState.WAITING_FOR_CLEAR)
    state.transition_to(AppState.PAUSED)
    state.transition_to(AppState.WAITING_FOR_CLEAR)
    assert state.state is AppState.WAITING_FOR_CLEAR


def test_invalid_transition_is_rejected() -> None:
    state = StateMachine()
    with pytest.raises(StateTransitionError):
        state.transition_to(AppState.COMPLETED)
