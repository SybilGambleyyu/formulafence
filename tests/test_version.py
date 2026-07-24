from __future__ import annotations

import re
from pathlib import Path

from formulafence import __version__


def test_runtime_version_matches_project_metadata() -> None:
    project = Path(__file__).resolve().parents[1] / "pyproject.toml"
    match = re.search(r'^version = "([^"]+)"$', project.read_text(), re.MULTILINE)

    assert match is not None
    assert __version__ == match.group(1)
