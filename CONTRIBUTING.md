# Contributing to DOCX Editor Skill

Thank you for your interest in contributing! This project follows high standards for token efficiency, performance, and code quality.

## Development Environment Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/omselkara/docx-editor.git
   cd docx-editor
   ```

2. **Install dependencies in editable mode:**
   ```bash
   pip install -e ".[dev]"
   ```

## Coding Standards

- **Type Hints:** All new functions must have comprehensive type hints. We aim for `mypy --strict` compatibility.
- **Error Handling:** Use the `scripts/docx_engine/errors.py` helpers instead of raw error strings.
  - `errors.err(module, action, reason)`
  - `errors.ok(message)`
- **Performance:** For read operations, prefer LXML fast-paths when dealing with large XML structures.
- **Formatting:** Use `ruff` for linting and formatting.

## Testing

We use `pytest` for all tests. Please ensure all tests pass before submitting a PR.

```bash
# Run all tests
pytest tests/

# Run tests with coverage
pytest tests/ --cov=scripts/docx_engine

# Run performance benchmarks
pytest tests/test_performance.py -v
```

## Pull Request Process

1. Create a new branch for your feature or bugfix.
2. Ensure your changes are covered by tests.
3. Update `README.md` or `references/` if you add new commands.
4. Run `ruff check scripts/` to ensure no linting issues.
5. Submit your PR!

## Repository Structure

- `scripts/docx_agent.py`: CLI entry point.
- `scripts/docx_engine/`: Core logic split into specialized modules.
- `tests/`: Comprehensive test suite.
- `references/`: Detailed command and workflow documentation.
