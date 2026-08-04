"""Shared pytest configuration for dbml-sharepoint tests."""

import json
import os
from pathlib import Path
from typing import TYPE_CHECKING, Any

import pytest
from _reachability import NOT_YET_REACHED, evaluate
from hypothesis import HealthCheck, settings

from dbml_sharepoint.analysis import conditions as _conditions
from dbml_sharepoint.analysis import findings as _findings

if TYPE_CHECKING:
    from collections.abc import Callable

# Hypothesis's default 200ms-per-example deadline is wall-clock, and this suite
# runs under `-n auto` with up to eight workers competing for the same cores.
# An example that takes 30ms alone can exceed 200ms while seven siblings run,
# so the deadline would fail on LOAD rather than on a slow property -- a flake
# that looks like a real failure and does not reproduce. Nothing here is being
# guarded by it; the properties are cheap.
settings.register_profile(
    "default",
    deadline=None,
    suppress_health_check=[HealthCheck.too_slow],
)

# CI searches harder than the local loop does. Measured on the property file
# alone, serially: 2.0s at the default 100 examples, 3.7s at 200, 8.1s at 500.
# The search space is small -- fourteen operators over shallow trees -- so 500
# buys little for four extra seconds. Under `-n auto` the whole suite goes from
# 6.05s to 6.63s with this profile, which is the number that actually matters.
settings.register_profile("ci", parent=settings.get_profile("default"), max_examples=200)

settings.load_profile("ci" if os.environ.get("CI") else "default")


# --- The per-finding reachability gate -------------------------------------
#
# See `_reachability.py` for why an aggregate coverage floor cannot do this job.
# The collection side lives here because it has to span the whole session, and
# under `-n auto` that means spanning processes too.

_FLAG = "--check-finding-reachability"
_SEEN: set[str] = set()


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        _FLAG,
        action="store_true",
        default=False,
        help=(
            "Fail the run if a declared FindingCode outside NOT_YET_REACHED is "
            "never constructed. Whole-suite only; CI passes it."
        ),
    )


def _record(name: str) -> None:
    _SEEN.add(name)


def _instrument() -> None:
    """Watch both carriers of a `FindingCode`.

    `Finding` is the usual one. `conditions._RefusalError` is the other: the
    condition compiler raises it with a code and its caller turns that into a
    `Finding` later, so watching `Finding` alone reports every condition rule as
    unreached. Measured -- instrumenting only `Finding` put 40 codes on the
    unreached list, 19 of them purely from this.
    """
    original_finding: Callable[..., None] = _findings.Finding.__init__
    original_refusal: Callable[..., None] = _conditions._RefusalError.__init__

    def finding_init(self: Any, *args: Any, **kwargs: Any) -> None:
        original_finding(self, *args, **kwargs)
        _record(self.code.name)

    def refusal_init(self: Any, code: Any, message: str) -> None:
        original_refusal(self, code, message)
        _record(code.name)

    _findings.Finding.__init__ = finding_init  # type: ignore[method-assign]
    _conditions._RefusalError.__init__ = refusal_init  # type: ignore[method-assign]


def _shared_dir(config: pytest.Config) -> Path:
    return Path(config.rootpath) / ".pytest_cache" / "finding-reach"


def _worker_id(config: pytest.Config) -> str:
    # xdist gives workers a `workerinput`; the controller has none, and neither
    # does a plain `-n0` run.
    return str(getattr(config, "workerinput", {}).get("workerid", "controller"))


def pytest_configure(config: pytest.Config) -> None:
    if not config.getoption(_FLAG):
        return
    _instrument()
    if _worker_id(config) == "controller":
        # Workers are spawned after this, so clearing here cannot race them.
        shared = _shared_dir(config)
        for stale in shared.glob("*.json"):
            stale.unlink()
        shared.mkdir(parents=True, exist_ok=True)


def pytest_sessionfinish(session: pytest.Session, exitstatus: int) -> None:
    config = session.config
    if not config.getoption(_FLAG):
        return

    shared = _shared_dir(config)
    shared.mkdir(parents=True, exist_ok=True)
    worker = _worker_id(config)
    if worker != "controller":
        # Same shape pytest-cov uses for the same problem: each worker drops its
        # slice, the controller combines once every worker has finished.
        (shared / f"{worker}.json").write_text(
            json.dumps(sorted(_SEEN)), encoding="utf-8",
        )
        return

    seen = set(_SEEN)
    for slice_file in shared.glob("*.json"):
        seen.update(json.loads(slice_file.read_text(encoding="utf-8")))

    if exitstatus != 0:
        # A failing suite reaches fewer codes for reasons that have nothing to
        # do with this gate. Reporting them all would bury the real failure.
        return

    problems = evaluate(
        declared={code.name for code in _findings.FindingCode},
        seen=seen,
        allowed_unreached=NOT_YET_REACHED,
    )
    if problems:
        session.exitstatus = 1
        reporter = config.pluginmanager.get_plugin("terminalreporter")
        if reporter is not None:
            reporter.write_sep("=", "finding reachability", red=True)
            for problem in problems:
                reporter.write_line(problem)
