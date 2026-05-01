"""Standardized error and success message builders.

Every public function should return strings from these helpers so tests
can assert on exact formats and callers can detect status reliably.

Status prefixes:
  SUCCESS: ...
  ERROR:   ...
  WARNING: ...
"""


def err(module: str, action: str, reason: str) -> str:
    return f"ERROR: {module}/{action}: {reason}"


def warn(module: str, action: str, reason: str) -> str:
    return f"WARNING: {module}/{action}: {reason}"


def ok(message: str) -> str:
    return f"SUCCESS: {message}"


def is_error(result: str) -> bool:
    return isinstance(result, str) and result.startswith("ERROR:")


def is_warning(result: str) -> bool:
    return isinstance(result, str) and result.startswith("WARNING:")


def is_success(result: str) -> bool:
    return isinstance(result, str) and result.startswith("SUCCESS:")
