import os

from build_windows import build_environment


def test_build_excludes_unrelated_dll_directories(monkeypatch):
    monkeypatch.setenv("PATH", r"C:\OtherTool\Poppler;C:\OtherTool\Qt")
    env = build_environment()
    assert "Poppler" not in env["PATH"]
    assert r"OtherTool\Qt" not in env["PATH"]
    assert "System32" in env["PATH"]
    assert "Poppler" in os.environ["PATH"]  # Parent environment is untouched.
