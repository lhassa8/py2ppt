# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and Development Commands

```bash
# Install development dependencies
pip install -e ".[dev]"

# Install pre-commit hooks
pre-commit install

# Run all tests
pytest tests/ -v

# Run tests with coverage
pytest tests/ -v --cov=py2ppt --cov-report=html

# Run a single test file
pytest tests/test_presentation.py -v

# Lint code
ruff check py2ppt/

# Format code
ruff format py2ppt/

# Type check
mypy py2ppt/
```

## Architecture

py2ppt is an AI-native PowerPoint library with a layered architecture:

### Layer Structure

1. **`py2ppt/oxml/`** - Low-level Open XML handling
   - Direct manipulation of PresentationML XML elements
   - `ns.py`: XML namespace definitions, `qn()` helper for Clark notation tags
   - `package.py`: ZIP package management for .pptx files
   - `slide.py`, `layout.py`, `master.py`, `theme.py`: XML operations for specific parts

2. **`py2ppt/core/`** - High-level abstractions
   - `presentation.py`: Presentation class wrapping the package
   - `slide.py`: Slide abstraction
   - `placeholder.py`: Placeholder management

3. **`py2ppt/tools/`** - Public API (tool functions for AI agents)
   - Each function is a discrete, stateless operation
   - `presentation.py`: create/open/save
   - `slides.py`: add/delete/duplicate/reorder
   - `content.py`: set_title/set_body/add_bullet
   - `media.py`: add_table/add_image
   - `inspection.py`: list_layouts/describe_slide/get_theme_colors
   - `style.py`: set_text_style

4. **`py2ppt/template/`** - Template analysis
   - `analyzer.py`: Extract template capabilities
   - `matcher.py`: Fuzzy layout name matching
   - `tokens.py`: Design token system for brand consistency

5. **`py2ppt/validation/`** - Style guide enforcement
   - `rules.py`: StyleGuide definition
   - `validator.py`: Validate presentations against rules

### Key Design Principles

- **Tool-First API**: Functions work as standalone tools for LLM agents (no method chaining)
- **Semantic Abstraction**: Use layout names ("Title Slide") not indices
- **Fuzzy Matching**: Layout names are fuzzy-matched (e.g., "title" finds "Title Slide")
- **Structured Errors**: Custom exceptions in `utils/errors.py` for AI self-correction

### Adding New Features

New tool functions go in `py2ppt/tools/`, then export via:
1. `py2ppt/tools/__init__.py`
2. `py2ppt/__init__.py`
