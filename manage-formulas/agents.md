---
name: Aspose.Cells Formula Management Agent
category: manage-formulas
product: Aspose.Cells for .NET
language: C#
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: Create, edit, copy, audit, and manage Excel formulas and named ranges in C#
primary_apis: [Cell.Formula, Cell.FormulaLocal, Cell.SetArrayFormula, Cell.SetSharedFormula, Name.RefersTo, Cell.GetPrecedents, Cell.GetDependents]
related_categories: [../calculate-formulas/, ../managing-ranges/, ../cells-data/, ../working-with-tables/]
---

# Formula Management Agent Instructions

## Mission and category boundary

Create focused examples for authoring, editing, copying, inspecting, auditing, and organizing Excel formulas with Aspose.Cells for .NET. Follow [`../AGENTS.md`](../AGENTS.md).

This category manages formula expressions and relationships. Use [`calculate-formulas`](../calculate-formulas/) when evaluation, calculation options, custom engines, monitoring, or calculation performance is the dominant intent.

## Scope

In scope: standard/local/R1C1 formulas, shared and array formulas, dynamic arrays, table calculated columns, named ranges, external references, precedents/dependents, formula search/replacement, auditing, validation, locking formula cells, and calculation-mode coordination after edits.

## Canonical API map

| Intent | APIs |
| --- | --- |
| Assign standard formula | `Cell.Formula` or verified `SetFormula` overload |
| Localized formula | `Cell.FormulaLocal` |
| R1C1 formula | Version-supported R1C1 property/method |
| Array formula | `SetArrayFormula` / dynamic-array APIs verified for the package |
| Shared formula | `SetSharedFormula` |
| Named range | `NameCollection`, `Name.RefersTo`, range APIs |
| Audit dependencies | `GetPrecedents`, `GetDependents`, calculation-specific dependency APIs as documented |
| Evaluate after editing | `Workbook.CalculateFormula` from the calculation workflow |

## Hard rules

- Begin formula strings with `=` and use documented invariant syntax unless `FormulaLocal` is explicitly required.
- Distinguish formula text from calculated value.
- Recalculate before asserting a result, but keep calculation mechanics secondary here.
- Use valid sheet quoting and absolute/relative references deliberately.
- Do not perform naive string replacement when formula token boundaries, quoted strings, or external links can change meaning.
- Verify supported functions and dynamic-array behavior for the installed package.
- Treat external workbook links and user-provided formulas as untrusted input.
- Do not label Excel functions "deprecated" without an authoritative source and compatibility rationale.

## Canonical pattern

```csharp
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];

worksheet.Cells["A1"].PutValue(10);
worksheet.Cells["A2"].PutValue(20);
worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

workbook.CalculateFormula();
double result = worksheet.Cells["A3"].DoubleValue;

if (Math.Abs(result - 30.0) > 0.000001)
{
    throw new InvalidOperationException("Formula result was not 30.");
}

workbook.Save("managed-formula.xlsx");
```

## Named ranges and dependencies

- Use stable, valid Excel names and avoid collisions with cell references.
- State workbook versus worksheet scope.
- Update `RefersTo` with valid formula syntax and verify dependent formulas after recalculation.
- Treat precedent/dependent results as potentially spanning worksheets/workbooks.
- Do not claim a complete dependency graph from an API that returns only direct relationships.

## Array, shared, and table formulas

- State target range dimensions and anchor cell.
- Do not overwrite spill ranges or table-managed columns accidentally.
- Verify every resulting cell or the expected spilled range.
- Use the API designed for the formula type instead of copying text manually.

## Example contract

Each example must identify formula type, target cell/range, expression, dependencies, primary API, expected formula text/value, and output. Use deterministic data and concise action-first filenames.

Metadata and the opening comment should answer one intent such as "set SUM formula in C#," "create named range formula," or "find formula precedents." Avoid keyword stuffing.

## Security and validation

- Validate user-supplied formulas against an allowlist/policy appropriate to the application.
- Audit external links, volatile functions, hidden-sheet references, and formula injection risks.
- Never fetch external workbooks or execute arbitrary code from formula text.
- Compile, run, verify stored formula text, calculate when required, assert values/dependencies, save/reopen, and confirm persistence.

Reject invented functions/APIs, unsafe text replacement, broken references, unsupported compatibility claims, and results read before calculation.

## Related knowledge

- [Category overview](README.md)
- [Calculate formulas](../calculate-formulas/)
- [Ranges](../managing-ranges/)
- [Tables](../working-with-tables/)
- [Official formula documentation](https://docs.aspose.com/cells/net/using-formulas-or-functions-to-process-data/)

## Definition of done

The example is done when formula type, syntax, target, dependencies, security assumptions, stored expression, calculated result where relevant, and persistence are explicit and verified.
