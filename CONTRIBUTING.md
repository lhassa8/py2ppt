# Contributing to py2ppt

Thank you for your interest in contributing to py2ppt! This document provides guidelines and information for contributors.

## Code of Conduct

By participating in this project, you agree to abide by our [Code of Conduct](CODE_OF_CONDUCT.md).

## Getting Started

### Development Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/lhassa8/py2ppt.git
   cd py2ppt
   ```

2. Create a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. Install development dependencies:
   ```bash
   pip install -e ".[dev]"
   ```

4. Install pre-commit hooks:
   ```bash
   pre-commit install
   ```

### Running Tests

```bash
# Run all tests
pytest tests/ -v

# Run with coverage
pytest tests/ -v --cov=py2ppt --cov-report=html

# Run specific test file
pytest tests/test_presentation.py -v
```

### Code Quality

We use several tools to maintain code quality:

- **Ruff**: Linting and formatting
- **MyPy**: Type checking
- **Pre-commit**: Automated checks before commits

```bash
# Run linter
ruff check py2ppt/

# Run formatter
ruff format py2ppt/

# Run type checker
mypy py2ppt/
```

## How to Contribute

### Reporting Bugs

1. Check if the bug has already been reported in [Issues](https://github.com/lhassa8/py2ppt/issues)
2. If not, create a new issue with:
   - A clear, descriptive title
   - Steps to reproduce the issue
   - Expected behavior vs actual behavior
   - Python version and OS
   - Minimal code example if possible

### Suggesting Features

1. Check existing issues and discussions for similar suggestions
2. Create a new issue with the "enhancement" label
3. Describe the feature and its use case
4. Explain why it would benefit py2ppt users

### Pull Requests

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature-name`
3. Make your changes
4. Add or update tests as needed
5. Ensure all tests pass: `pytest tests/ -v`
6. Ensure code quality checks pass: `ruff check py2ppt/ && mypy py2ppt/`
7. Commit with a clear message
8. Push to your fork
9. Open a Pull Request

#### Pull Request Guidelines

- Keep PRs focused on a single change
- Include tests for new functionality
- Update documentation as needed
- Follow existing code style
- Add type hints to all public functions
- Write clear commit messages

## Project Structure

```
py2ppt/
├── oxml/           # Low-level Open XML handling
│   ├── package.py  # ZIP package management
│   ├── slide.py    # Slide XML operations
│   └── ...
├── core/           # High-level abstractions
│   ├── presentation.py
│   ├── slide.py
│   └── ...
├── tools/          # Public API (tool functions)
│   ├── presentation.py
│   ├── slides.py
│   ├── content.py
│   └── ...
├── template/       # Template analysis
├── validation/     # Style guide validation
└── utils/          # Utilities
```

### Key Design Principles

1. **Tool-First API**: Every public function should work as a standalone tool for AI agents
2. **Smart Defaults**: Functions should work without explicit template analysis
3. **Structured Errors**: Return clear, actionable error information
4. **Type Safety**: All public APIs should have complete type hints
5. **Minimal Dependencies**: Keep the core dependency-light

## Adding New Features

### Adding a New Tool Function

1. Add the function to the appropriate file in `py2ppt/tools/`
2. Export it in `py2ppt/tools/__init__.py`
3. Export it in `py2ppt/__init__.py`
4. Add comprehensive tests
5. Add docstring with clear parameters and return value

Example:
```python
def my_new_tool(
    presentation: Presentation,
    param1: str,
    *,
    optional_param: int = 10,
) -> dict:
    """Brief description of what this tool does.

    Args:
        presentation: The presentation to modify
        param1: Description of param1
        optional_param: Description with default value

    Returns:
        Description of return value

    Raises:
        SpecificError: When this error occurs
    """
    ...
```

### Adding Open XML Support

When adding new Open XML functionality:

1. Add XML handling to the appropriate file in `py2ppt/oxml/`
2. Use proper namespaces from `py2ppt/oxml/ns.py`
3. Test with real .pptx files opened in PowerPoint/LibreOffice
4. Ensure round-trip compatibility (open → modify → save → reopen)

## Questions?

Feel free to open an issue for questions or join discussions in the repository.

Thank you for contributing!
