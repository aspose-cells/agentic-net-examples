---
name: Aspose.Cells VBA and Macro Project Agent
category: macro-project
product: Aspose.Cells for .NET
language: C#
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Inspect, create, modify, copy, protect, and sign Excel VBA projects in C#
primary_apis: [VbaProject, VbaModule, VbaModuleCollection, VbaProjectReferenceCollection]
related_categories: [../encryption-and-protection/, ../open-workbook/, ../save-workbook/]
---

# VBA and Macro Project Agent Instructions

## Mission and security posture

Create focused Aspose.Cells for .NET examples for VBA project inspection and preservation, module editing, references, protection, and digital signatures. Follow [`../AGENTS.md`](../AGENTS.md).

Aspose.Cells manages VBA project content; examples must never claim that macros are executed by Aspose.Cells. Treat all VBA code and references as untrusted active content.

## Scope

In scope: `Workbook.VbaProject`, modules, code, module names/types, copying projects/modules, COM references, designer storage, protection, signatures, validation, macro assignment, and macro-enabled save formats.

Keep general workbook signatures in `encryption-and-protection` and ordinary workbook saving in `save-workbook`.

## Hard rules

- Save macro-bearing workbooks in macro-capable formats such as XLSM or XLSB as appropriate.
- Never save to XLSX while claiming VBA is preserved.
- Do not execute, compile, or trust imported VBA code.
- Do not add system COM references unless the task requires them and the environment dependency is explicit.
- Use test certificates only; never embed private-key passwords or real certificates.
- Distinguish a valid signature from an unchanged signature: modifying VBA normally invalidates/removes prior trust.
- Preserve module attributes and designer storage when copying them is the intent.
- Verify `VbaProject.IsValidSigned` or current equivalent only after reopening where signature persistence matters.

## Canonical inspection pattern

```csharp
Workbook workbook = new Workbook("input.xlsm");
VbaProject project = workbook.VbaProject;

Console.WriteLine($"Module count: {project.Modules.Count}");
foreach (VbaModule module in project.Modules)
{
    Console.WriteLine(module.Name);
}

workbook.Save("macro-project-copy.xlsm", SaveFormat.Xlsm);
```

## Example contract

Each example must state whether it inspects, modifies, copies, protects, or signs VBA; identify input/output macro format; use synthetic harmless VBA; and verify module/reference/signature state.

Metadata must include macro operation, primary API, trust assumptions, output format, and expected result. Prefer filenames such as `list-vba-modules-in-xlsm.cs`.

## Enterprise safety and validation

- Analyze or display VBA as text only; never invoke it.
- Avoid logging complete production macro source.
- Reject external reference paths from untrusted inputs.
- Keep certificate access platform-aware and CI-safe.
- Save/reopen and compare module count, names, code hashes or reference metadata as appropriate.
- Confirm that non-macro formats intentionally remove VBA when removal is the scenario.

Reject examples that use nonexistent module constructors, assume COM libraries are installed, expose secrets, or claim signature validity after project modification without verification.

## SEO, GEO, and AEO

Target one direct intent such as "read VBA modules from XLSM in C#," "copy Excel macros," or "sign a VBA project." State clearly that Aspose.Cells manages but does not execute VBA.

## Related knowledge

- [Category overview](README.md)
- [Encryption and protection](../encryption-and-protection/)
- [Open workbook](../open-workbook/)
- [Official VBA documentation](https://docs.aspose.com/cells/net/manage-vba-project/)

## Definition of done

The example is done when active-content risk, macro format, operation, environment dependencies, and post-save project/signature state are explicit and verified.

