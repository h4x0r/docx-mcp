# Security Hardening Design

**Date:** 2026-05-14  
**Status:** Approved  
**Threat model:** Untrusted user uploads — adversarial DOCX files and adversarial tool parameters

---

## Goal

Harden docx-mcp against malicious and crafted inputs across two surfaces:
1. **Document layer** — crafted DOCX files (ZipSlip, XXE, ZIP bomb, corrupt inputs)
2. **Parameter layer** — adversarial tool parameters (path traversal, ReDoS, unbounded integers)

Then add a fuzzing harness to discover unknown vulnerabilities continuously.

## Architecture

Three new components slot into the existing mixin architecture without restructuring:

1. **`docx_mcp/document/guards.py`** — centralized `InputGuard` class; all parameter validation lives here
2. **Fixes in existing files** — `base.py` (open/save), `reading.py` (search_text), `query.py` (xpath_query)
3. **`tests/fuzz/fuzz_open.py`** — `atheris` byte-level fuzzer harness

`server.py` calls `InputGuard.*` at the top of each tool handler. Document-layer fixes are self-contained.

## Tech Stack

- `hypothesis` — property-based testing (already a dev dependency or easy add)
- `atheris` — Google's Python fuzzing library (libFuzzer-backed); separate install
- `lxml` — already used; fixes are parser flag changes only
- `re`, `signal` — stdlib; used for regex timeout

---

## Vulnerability Inventory

| ID | Vulnerability | Location | Severity |
|----|--------------|----------|----------|
| V1 | ZipSlip — ZIP entry names with `../` escape workdir | `base.py open()` | Critical |
| V2 | XXE — XML parser resolves external entities | `base.py open()` | High |
| V3 | ZIP bomb — no size/entry-count limit on extraction | `base.py open()` | High |
| V4 | Unhandled `BadZipFile` / corrupt XML crash | `base.py open()` | High |
| V5 | ReDoS — user regex passed directly to `re.finditer` | `reading.py search_text()` | Medium |
| V6 | Output path traversal — arbitrary write via `output_path` | `base.py save()`, `reading.py copy_document()` | High |
| V7 | XPath DoS — unbounded complexity user XPath | `query.py xpath_query()` | Medium |

---

## Component Design

### `InputGuard` (`docx_mcp/document/guards.py`)

```python
class InputGuard:
    PARA_ID_RE   = re.compile(r'^[0-9A-Fa-f]{1,8}$')
    HEX_COLOR_RE = re.compile(r'^[0-9A-Fa-f]{6}$')
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100 MB

    @staticmethod
    def para_id(value: str) -> str:
        """Validate 1–8 hex digit para ID. Raises ValueError."""

    @staticmethod
    def output_path(value: str) -> Path:
        """Reject traversal, non-.docx suffix, symlinks outside cwd."""

    @staticmethod
    def input_path(value: str) -> Path:
        """Reject non-existent, non-.docx, oversized files."""

    @staticmethod
    def color_hex(value: str) -> str:
        """Validate 6-digit hex color string."""

    @staticmethod
    def bounded_int(value: int, lo: int, hi: int, name: str) -> int:
        """Reject integers outside [lo, hi]."""

    @staticmethod
    def regex_pattern(value: str) -> re.Pattern:
        """Compile regex; reject on compile error."""
```

Guards raise `ValueError` with a descriptive message. The MCP framework converts these to structured tool error responses.

### Document-layer fixes

**V1 — ZipSlip (`base.py`)**

Before `extractall`, validate every entry:
```python
for member in zf.namelist():
    dest = (workdir / member).resolve()
    if not str(dest).startswith(str(workdir)):
        raise DocxMcpError(ErrCode.UNSAFE_PATH, f"ZipSlip: {member}")
zf.extractall(workdir)
```

**V2 — XXE (`base.py`)**

Replace bare `etree.XMLParser(remove_blank_text=False)` with:
```python
etree.XMLParser(
    remove_blank_text=False,
    resolve_entities=False,
    no_network=True,
    huge_tree=False,
)
```

**V3 — ZIP bomb (`base.py`)**

Before extraction, sum uncompressed sizes and count entries:
```python
MAX_UNCOMPRESSED = 500 * 1024 * 1024  # 500 MB
MAX_ENTRIES = 10_000
infos = zf.infolist()
if len(infos) > MAX_ENTRIES:
    raise DocxMcpError(ErrCode.OOXML_INVALID, "Too many ZIP entries")
total = sum(i.file_size for i in infos)
if total > MAX_UNCOMPRESSED:
    raise DocxMcpError(ErrCode.OOXML_INVALID, "ZIP uncompressed size exceeds limit")
```

**V4 — BadZipFile (`base.py`)**

Wrap the `zipfile.ZipFile(...)` call:
```python
try:
    with zipfile.ZipFile(self.source_path, "r") as zf:
        ...
except zipfile.BadZipFile as exc:
    raise DocxMcpError(ErrCode.OOXML_INVALID, f"Corrupt DOCX: {exc}") from exc
```

**V5 — ReDoS (`reading.py`)**

Use `re.compile()` with a `signal.alarm` timeout:
```python
import signal

def _compile_regex(pattern: str, timeout: int = 2) -> re.Pattern:
    def _handler(signum, frame):
        raise TimeoutError
    old = signal.signal(signal.SIGALRM, _handler)
    signal.alarm(timeout)
    try:
        compiled = re.compile(pattern)
        # test on empty string to trigger catastrophic backtracking early
        compiled.search("")
    except TimeoutError:
        raise DocxMcpError(ErrCode.MALFORMED_INPUT, "Regex timed out (possible ReDoS)")
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, old)
    return compiled
```

Note: `signal.SIGALRM` is Unix-only. On Windows, use a thread-based timeout.

**V6 — Output path traversal (`base.py save()`, `reading.py copy_document()`)**

Call `InputGuard.output_path(output_path)` at the top of both methods.

**V7 — XPath DoS (`query.py`)**

Wrap `tree.xpath()` in the same `signal.alarm` pattern as V5 with a 2 s timeout.

### New error codes (`errors.py`)

```python
MALFORMED_INPUT = "MALFORMED_INPUT"   # bad parameter value
UNSAFE_PATH     = "UNSAFE_PATH"       # path traversal / unsafe suffix
```

---

## Fuzzing Strategy

### Layer A — hypothesis property-based tests

File: `tests/test_security_guards.py`

Uses `@given` strategies to exhaustively test `InputGuard.*`:
- `para_id`: arbitrary unicode, null bytes, very long strings, format strings
- `output_path`: traversal payloads, non-.docx extensions, symlinks
- `color_hex`: wrong lengths, non-hex chars, injection strings
- `bounded_int`: extremes (`sys.maxsize`, `-(2**63)`, zero)

Runs in CI with normal `pytest` — no extra tooling.

### Layer B — atheris byte-level fuzzer

File: `tests/fuzz/fuzz_open.py`

```python
import atheris, sys
from docx_mcp.document import DocxDocument
from docx_mcp.document.errors import DocxMcpError

def TestOneInput(data):
    tmp = write_to_tmp(data)
    try:
        doc = DocxDocument(tmp)
        doc.open()
        doc.close()
    except (DocxMcpError, FileNotFoundError, ValueError, OSError):
        pass  # expected
    except Exception as e:
        raise  # unexpected = fuzzer finds a crash

atheris.Setup(sys.argv, TestOneInput)
atheris.Fuzz()
```

The invariant: **any input must either succeed or raise one of the expected exceptions — never crash with an unhandled exception.**

Corpus seeds (`tests/fuzz/corpus/`): `minimal.docx`, `with_image.docx`, `with_table.docx`

Run locally: `python tests/fuzz/fuzz_open.py -runs=50000`  
CI smoke: `python tests/fuzz/fuzz_open.py -runs=1000`

---

## Test File Layout

```
tests/
  test_security_guards.py      # hypothesis + unit: InputGuard.*
  test_security_document.py    # unit: V1 ZipSlip, V2 XXE, V3 ZIP bomb, V4 BadZipFile
  test_security_search.py      # unit: V5 ReDoS timeout
  test_security_xpath.py       # unit: V7 XPath timeout
  fuzz/
    fuzz_open.py               # atheris harness
    corpus/
      minimal.docx
      with_image.docx
      with_table.docx
```

Each file maps to one vulnerability group. TDD: RED commit (failing tests) then GREEN commit (implementation) per task.

---

## Implementation Tasks (for writing-plans)

1. Add `MALFORMED_INPUT` and `UNSAFE_PATH` error codes to `errors.py`
2. Implement `InputGuard` in `guards.py` with unit tests (hypothesis)
3. Fix V4 — `BadZipFile` handling in `base.py`
4. Fix V1 — ZipSlip in `base.py`
5. Fix V2 — XXE parser flags in `base.py`
6. Fix V3 — ZIP bomb limits in `base.py`
7. Fix V5 — ReDoS timeout in `reading.py`
8. Fix V6 — output path guard in `base.py save()` and `reading.py copy_document()`
9. Fix V7 — XPath timeout in `query.py`
10. Build atheris fuzzer harness + corpus in `tests/fuzz/`
