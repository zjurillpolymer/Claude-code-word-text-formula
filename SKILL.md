---
name: docx-formula
description: Create Word .docx files with LaTeX formulas rendered as native OMML (Word's equation objects). Use whenever the user asks to write a .docx file containing mathematical formulas — especially for scientific/engineering writing. TRIGGER when: user mentions "docx" + "formula"/"equation"/"LaTeX in Word"; or any docx task that involves math expressions. Combine with the general `docx` skill for full document creation.
---

# Docx Formula Skill — LaTeX → Native Word Equations

Convert LaTeX math expressions into native Word OMML (Office Math Markup Language) equations in `.docx` files using `pandoc` as the conversion backend. Produces equations that Word can open, edit, and render natively — no images, no plain text.

## Requirements

- **pandoc ≥ 3.0** — `brew install pandoc` or `conda install -c conda-forge pandoc`
- **python-docx** — `pip install python-docx`
- **lxml** — `pip install lxml`

## Quick start

```python
from docx_math import add_formula, add_formula_to_document
from docx import Document

doc = Document()

# Inline formula after some text
p = doc.add_paragraph()
p.add_run('Gibbs free energy: ')
add_formula(p, r'\Delta G = \Delta H - T \Delta S')

# Display formula (centered, own line)
add_formula(doc.add_paragraph(), r'\frac{d[M]}{dt} = k_p [M]', display=True)

doc.save('output.docx')
```

## How it works

1. Formulas are batched and sent to `pandoc -f markdown -t docx`
2. Pandoc generates standard OMML (`m:oMath` / `m:oMathPara`) — the same format Word itself produces
3. OMML elements are extracted and inserted as direct children of `<w:p>` (standard OOXML placement)
4. Identical formulas are cached per session to avoid redundant pandoc calls

## API

### `add_formula(paragraph, latex, display=False)`

Render LaTeX as OMML and append it to *paragraph*.

| Parameter | Type | Default | Description |
|---|---|---|---|
| `paragraph` | `Paragraph` | — | python-docx paragraph to append into |
| `latex` | `str` | — | LaTeX math (without `$` delimiters) |
| `display` | `bool` | `False` | `True` → centered display formula (`m:oMathPara`) |

### `add_formula_to_document(doc, latex, display=True)`

Convenience: create a new paragraph containing a formula. Returns the paragraph.

## Supported LaTeX

Anything pandoc supports, including:

- **Fractions**: `\frac{a}{b}`
- **Sub/superscripts**: `x_i`, `x^2`, `x_i^2`
- **Greek letters**: `\alpha`, `\Delta`, `\sigma`
- **Sums & integrals**: `\sum_{i=1}^{N}`, `\int_{-\infty}^{\infty}`
- **Radicals**: `\sqrt{x}`, `\sqrt[n]{x}`
- **Operators**: `\mathbf{r}`, `\langle`, `\rangle`
- **Matrices**: `\begin{matrix} ... \end{matrix}`
- **Functions**: `\exp`, `\sin`, `\log`

## Inline vs display

| Mode | LaTeX | OMML element | Appearance |
|---|---|---|---|
| Inline | `\Delta G = ...` | `m:oMath` | Sits within text |
| Display | `\frac{1}{N} \sum ...` | `m:oMathPara` > `m:oMath` | Centered, own line |

## Mixing text and formulas

```python
p = doc.add_paragraph()
p.add_run('The Flory exponent satisfies: ')
add_formula(p, r'\langle R^2 \rangle^{1/2} \sim N^{\nu}')
p.add_run(', where \nu \approx 0.588 in good solvent.')
```

## File structure

```
Claude-code-word-text-formula/
├── SKILL.md              ← skill definition (auto-discovered by Claude Code)
├── docx_math.py          ← core Python module
├── README.md             ← this file
├── LICENSE
└── requirements.txt      ← python-docx, lxml
```
