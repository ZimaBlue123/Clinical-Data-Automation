# Logging Convention

This document defines a lightweight, grep-friendly logging style for all CLI modules in this repository.

## Core Principles

- Keep one event per log line.
- Use explicit key-value fields in message text.
- Prefer stable keys (`action`, `file`, `input`, `output`, `reason`, `result`).
- Include enough context to debug without reproducing immediately.
- Use `logger.exception(...)` only at true failure boundaries.

## Level Mapping

- `INFO`: normal progress, start/end events, summary counts.
- `WARNING`: recoverable issue, skip, fallback, degraded path.
- `ERROR`: terminal error for current command/input.
- `EXCEPTION`: unexpected exception with traceback (same as `ERROR` + stack).
- `DEBUG`: verbose internals for troubleshooting.

## Message Format

Recommended pattern:

`action=<event> key1=<value> key2=<value> ...`

Examples:

- `action=process_start file=sample.pdf`
- `action=process_success file=sample.pdf output=output/sample_clean.pdf`
- `action=skip_existing file=sample.pdf reason=output_exists`
- `action=download_failed query=10.1000/xyz reason=timeout`
- `action=batch_complete success=18 failed=2 total=20`

## Required Context by Scenario

- File processing: include `file`, and where relevant `output`.
- Rule engine: include `rule`, `keyword`, `page`.
- Batch execution: include `success`, `failed`, `total`.
- Input validation: include `input` and failure `reason`.
- External I/O (network/API): include request target/query identifier and `reason`.

## Exception Handling Rules

- Use `logger.exception("action=... ...")` in top-level task boundaries only.
- Avoid swallowing exceptions silently (`except: pass`) unless explicitly safe.
- For safe fallback blocks, log at `DEBUG` or `WARNING` with rationale.

## CLI Exit Code Convention

- `0`: success (or fully successful batch).
- `1`: operational failure (invalid input, partial/complete processing failure).
- `2`: dependency/preflight failure (optional, when explicitly used by module).
- `130`: user interrupted (`KeyboardInterrupt`).

## Python Snippet Template

```python
import logging

logger = logging.getLogger(__name__)

def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    logger.info("action=batch_start input=%s", input_path)
    try:
        # do work...
        logger.info("action=batch_complete success=%s failed=%s total=%s", ok, failed, total)
        return 0 if failed == 0 else 1
    except KeyboardInterrupt:
        logger.error("action=interrupted_by_user")
        return 130
    except Exception:
        logger.exception("action=batch_failed")
        return 1
```

## Adoption Notes

- For legacy modules, prioritize standardizing top-level CLI entry logs first.
- Preserve user-facing behavior; improve logs without changing business logic unless needed.
- Keep wording consistent to support future log aggregation and automated diagnostics.
