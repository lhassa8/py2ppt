# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and Development Commands

```bash
# Install development dependencies
pip install -e ".[dev]"

# Install pre-commit hooks
pre-commit install

# Run all tests
pytest tests/test_py2ppt/ -v

# Run tests with coverage
pytest tests/test_py2ppt/ -v --cov=py2ppt --cov-report=html

# Run a single test file
pytest tests/test_py2ppt/test_template.py -v

# Lint code
ruff check py2ppt/

# Format code
ruff format py2ppt/

# Type check
mypy py2ppt/
```

## Architecture

py2ppt is an AI-friendly wrapper around **python-pptx** for creating PowerPoint presentations from corporate templates with semantic, intent-based APIs.

### Package Structure

- **`py2ppt/template.py`** - `Template` class: loads a .pptx template, analyzes its layouts, recommends layouts by content type, and creates new presentations
- **`py2ppt/presentation.py`** - `Presentation` class: high-level slide creation (`add_title_slide`, `add_content_slide`, `add_comparison_slide`, etc.) and save
- **`py2ppt/layout.py`** - Layout analysis: `LayoutType` enum, `classify_layout`, `analyze_layout`, `recommend_layout`
- **`py2ppt/placeholders.py`** - Semantic placeholder mapping: `PlaceholderRole`, `map_placeholders`, `get_placeholder_purpose`
- **`py2ppt/formatting.py`** - Bullet/text formatting: `parse_content`, `auto_bullets`, `FormattedParagraph`, `FormattedRun`

### Key Design Principles

- **AI-Native API**: Designed for LLM agents to discover layouts, get recommendations, and create slides semantically
- **Template-First**: Always starts from a corporate .pptx template — analyzes its layouts, colors, and fonts
- **Semantic Abstraction**: Use content types ("title", "content", "comparison") not layout indices
- **Wraps python-pptx**: All underlying PowerPoint manipulation uses python-pptx; py2ppt adds the semantic layer on top

### Tests

Tests live in `tests/test_py2ppt/`. Some tests require the `AWStempate.pptx` file in the repo root and will skip if it's not present.
