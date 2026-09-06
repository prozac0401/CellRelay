"""Regression inputs transcribed from UI labels only; no screenshot user data."""

import pytest

from app.browser import aws_labels as labels
from app.browser.aws_skill_builder import AwsSkillBuilderAssignment


@pytest.mark.parametrize(
    "pattern,text",
    [
        (labels.TRAINING_ASSIGN, "Assign training"),
        (labels.ASSIGN_TO_USER, "Assign to users"),
        (labels.USER_DIALOG, "Select users"),
        (labels.USER_SEARCH, "Find users"),
        (labels.CANCEL, "Cancel"),
        (labels.ASSIGN, "Assign"),
        (labels.NO_MATCHES, "No matches"),
        (labels.NO_MATCHES, "No data found"),
        (labels.REGISTER_ALL, "I want to enroll all selected users."),
        (labels.DONE, "Done"),
    ],
)
def test_observed_english_labels(pattern, text):
    assert pattern.fullmatch(text)


@pytest.mark.parametrize(
    "query,email,expected",
    [
        ("learner@", "learner@example.com", True),
        (" LEARNER@ ", "learner@example.com", True),
        ("learner@", "other.learner@example.com", False),
        ("learner@", "learner2@example.com", False),
        ("learner@exam", "learner@example.com", False),
        ("learner@@", "learner@example.com", False),
        ("@", "learner@example.com", False),
        ("learner@example.com", "learner@other.example", False),
    ],
)
def test_trailing_at_matches_only_a_complete_local_part(query, email, expected):
    assert AwsSkillBuilderAssignment.matches_user(query, email) is expected
