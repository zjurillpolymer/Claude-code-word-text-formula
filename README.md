# docx-formula — LaTeX to Native Word Equations

A [Claude Code](https://docs.anthropic.com/en/docs/claude-code) skill that converts LaTeX math expressions into native Word OMML equations in `.docx` files. Uses `pandoc` as the conversion backend — produces equations that Word can open, edit, and render natively.

## Why this exists

When creating Word documents programmatically, writing LaTeX formulas as plain text doesn't work — Word can't render them. And manually constructing OMath XML is error-prone and fragile across Word versions.

This skill takes a different approach: **use pandoc to generate the OMML**. Pandoc produces the same format Word itself uses (`m:oMath` / `m:oMathPara`), so the output is always valid and editable.

## Requirements

| Tool | Install |
|---|---|
| **pandoc ≥ 3.0** | `brew install pandoc` or `conda install -c conda-forge pandoc` |
| **python-docx** | `pip install python-docx` |
| **lxml** | `pip install lxml` |

## Install as a Claude Code skill

```bash
git clone https://github.com/zjurillpolymer/Claude-code-word-text-formula.git \
  ~/.claude/skills/docx-formula

# Or if you have the SSH URL:
git clone git@github.com:zjurillpolymer/Claude-code-word-text-formula.git \
  ~/.claude/skills/docx-formula
```

After cloning, `docx-formula` appears in Claude Code's available skills. It auto-triggers when you mention creating a `.docx` with math formulas.

## Usage

```python
from docx_math import add_formula, add_formula_to_document
from docx import Document

doc = Document()

# Inline formula
p = doc.add_paragraph()
p.add_run('Gibbs free energy: ')
add_formula(p, r'\Delta G = \Delta H - T \Delta S')

# Display formula (centered, own line)
add_formula(doc.add_paragraph(), r'\frac{d[M]}{dt} = k_p [M]', display=True)

doc.save('output.docx')
```

## API

### `add_formula(paragraph, latex, display=False)`

Render LaTeX as OMML and append to *paragraph*.

| Param | Type | Default | Description |
|---|---|---|---|
| `paragraph` | `Paragraph` | — | python-docx paragraph to append into |
| `latex` | `str` | — | LaTeX math (without `$` delimiters) |
| `display` | `bool` | `False` | `True` → centered display formula |

### `add_formula_to_document(doc, latex, display=True)`

Convenience: new paragraph with formula. Returns the paragraph.

## How it works

```
LaTeX formula
    → pandoc -f markdown -t docx
    → docx with native OMML
    → extract m:oMath / m:oMathPara elements
    → insert into target docx as <w:p> direct children
```

- Pandoc handles all LaTeX parsing and OMML generation
- OMML elements are placed as direct children of `<w:p>` (standard OOXML), not inside `<w:r>`
- Identical formulas are cached per session

## Supported LaTeX

Fractions, sub/superscripts, Greek letters, sums/integrals with limits, radicals, matrices, operators like `\mathbf`, `\langle`/`\rangle`, `\exp`, `\sin`, etc. — anything pandoc supports.

## License

MIT
