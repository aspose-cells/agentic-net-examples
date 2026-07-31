---
name: Aspose.Cells Cell Formatting Agent
category: format-cells
product: Aspose.Cells for .NET
language: C#
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: Format Excel cells, rows, columns, and ranges in C#
primary_apis: [Style, Cell.GetStyle, Cell.SetStyle, Workbook.CreateStyle, StyleFlag, Cells.ApplyStyle]
related_categories: [../cells-data/, ../managing-ranges/, ../rows-and-columns/, ../globalization-and-localization/]
---

# Cell Formatting Agent Instructions

## Mission and scope

Create focused Aspose.Cells for .NET examples for fonts, fills, borders, alignment, number/date formats, text wrapping, themes, gradients, style reuse, and range/row/column formatting. Follow [`../AGENTS.md`](../AGENTS.md).

Use `globalization-and-localization` when culture behavior is primary and `cells-data` when values rather than presentation are primary.

## Canonical API map

| Goal | API |
| --- | --- |
| Modify one cell style | `Cell.GetStyle` then `Cell.SetStyle` |
| Create reusable style | `Workbook.CreateStyle` |
| Apply selected attributes | `StyleFlag` with `ApplyStyle` |
| Apply to row/column | Verified row/column style methods on `Cells` |
| Number/date display | `Style.Number` or `Style.Custom` |
| Fill | `Style.ForegroundColor`, `BackgroundColor`, `Pattern` |
| Border | `Style.Borders` and `BorderType` |
| Alignment | `Style.HorizontalAlignment`, `VerticalAlignment`, wrapping/rotation properties |

## Hard rules

- Treat `GetStyle()` as a style value to modify and reapply with `SetStyle()`.
- Set a fill pattern when a foreground/background color requires it.
- Use `StyleFlag` so bulk operations change only intended attributes.
- Reuse styles for large regions; avoid creating a unique style per cell.
- Keep value and display format distinct; number formatting does not change the underlying value.
- Use theme colors only when theme behavior is the intent; otherwise deterministic RGB colors are clearer.
- Verify formatting after save/reopen and, for visual claims, render a representative output.

## Canonical pattern

```csharp
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
Cell cell = worksheet.Cells["A1"];
cell.PutValue(1250.5);

Style style = cell.GetStyle();
style.Custom = "$#,##0.00";
style.Font.IsBold = true;
cell.SetStyle(style);

workbook.Save("formatted-cell.xlsx");
```

## Example contract

Each example must demonstrate one dominant formatting outcome, use meaningful sample data, identify the target cell/range, verify style properties, and save/reopen when persistence matters. Metadata should name the style feature, primary API, target, output, and expected appearance/property.

For visual comparisons, specify RGB/theme values, format codes, border side/style, alignment, and render conditions. Never claim pixel-perfect rendering across environments without accounting for fonts and platform differences.

## Scale, accessibility, and safety

- Apply styles to ranges rather than per-cell loops where possible.
- Avoid style explosion; report style counts in performance examples.
- Maintain readable contrast and do not use color as the sole carrier of meaning.
- Sanitize user-provided custom number formats and do not treat display strings as trusted numeric values.
- Avoid developer-specific fonts unless fallback behavior is the scenario.

## Discoverability and validation

Target intents such as "format Excel cells in C#," "apply number format," or "add borders with Aspose.Cells." The opening comment must state the target and visible/property result.

Verify exact style properties and enums against the installed package. Compile, run, reopen, and compare style fields; render when the claim is visual. Reject code that modifies a style without reapplying it or accidentally overwrites unrelated attributes.

## Related knowledge

- [Category overview](README.md)
- [Cell data](../cells-data/)
- [Ranges](../managing-ranges/)
- [Localization](../globalization-and-localization/)
- [Official formatting documentation](https://docs.aspose.com/cells/net/cells-formatting/)

## Definition of done

The example is done when target, style attributes, underlying value, expected appearance, persistence, and performance implications are explicit and verified.
