---
name: Aspose.Cells Globalization and Localization Agent
category: globalization-and-localization
product: Aspose.Cells for .NET
language: C#
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Process localized Excel dates, numbers, formulas, labels, and regional settings in C#
primary_apis: [GlobalizationSettings, CultureInfo, WorkbookSettings, LoadOptions, Cell.FormulaLocal]
related_categories: [../format-cells/, ../manage-formulas/, ../conversion/, ../cells-data/]
---

# Globalization and Localization Agent Instructions

## Mission and scope

Create deterministic, culture-aware Aspose.Cells for .NET examples for localized numbers, dates, calendars, formulas, Boolean/error labels, subtotal/pivot labels, chart text, and regional workbook behavior. Follow [`../AGENTS.md`](../AGENTS.md).

Use `format-cells` when only a format string changes, and `manage-formulas` when formula structure rather than language/culture is primary.

## Canonical model

| Concern | Preferred mechanism |
| --- | --- |
| Parse/load with a culture | Version-supported `LoadOptions` culture/region settings |
| Workbook region | `Workbook.Settings.Region` or package-supported culture setting |
| Custom labels/function names | Derive from `GlobalizationSettings` and override documented members |
| Localized formula text | `Cell.FormulaLocal` when appropriate and supported |
| Culture-specific display | Typed values plus `Style.Custom`/number formats |
| Japanese era/calendar | Explicit `CultureInfo`/calendar APIs and verified Aspose helpers |

## Hard rules

- Set culture explicitly; never rely on the developer machine's current culture.
- Store dates and numbers as typed values, not preformatted localized strings, unless parsing text is the objective.
- Keep invariant formula syntax and localized formula syntax distinct.
- Save and restore thread culture if an example changes it.
- Use valid BCP 47/.NET culture identifiers and deterministic dates away from ambiguous boundaries unless testing boundaries.
- Explain whether localization affects stored values, display formats, formula names, or rendered labels.
- Do not claim that changing culture translates arbitrary workbook text.

## Canonical culture-safe pattern

```csharp
CultureInfo culture = CultureInfo.GetCultureInfo("fr-FR");
Workbook workbook = new Workbook();
workbook.Settings.CultureInfo = culture;
Worksheet worksheet = workbook.Worksheets[0];

worksheet.Cells["A1"].PutValue(new DateTime(2026, 6, 29));
Style style = worksheet.Cells["A1"].GetStyle();
style.Custom = "dd mmmm yyyy";
worksheet.Cells["A1"].SetStyle(style);

workbook.Save("localized-workbook.xlsx");
Console.WriteLine(culture.Name);
```

When culture must influence workbook parsing or rendering, configure the documented workbook/load setting rather than creating an unused `CultureInfo` object.

## Custom globalization settings

- Override only documented methods needed by the scenario.
- Provide a deterministic fallback for unknown labels or functions.
- Keep mappings case-aware according to Excel semantics.
- Never perform network calls or workbook mutation inside localization callbacks.
- Unit-test each locale mapping and fallback.

## Example contract

Each example must identify locale, input representation, stored type, display/formula behavior, output, and expected localized result. Metadata and opening comments must use canonical locale identifiers. Prefer filenames such as `load-excel-with-french-culture.cs`.

## Validation and discoverability

Test under an explicitly configured culture, compare typed values separately from display text, save/reopen, and render when visual labels are claimed. Restore global/thread state in `finally`.

Target one intent such as "format Excel dates for French culture" or "localize Excel formula names in C#." Reject examples that depend silently on installed Excel language packs, compare culture-sensitive strings without a declared culture, or alter process-wide culture without cleanup.

## Related knowledge

- [Category overview](README.md)
- [Cell formatting](../format-cells/)
- [Formula management](../manage-formulas/)
- [Official globalization documentation](https://docs.aspose.com/cells/net/globalization-and-localization/)

## Definition of done

The example is done when locale, stored value, display/formula effect, fallback behavior, process-state cleanup, and expected output are explicit and verified.

