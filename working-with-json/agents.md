---
name: Aspose.Cells JSON Agent
category: working-with-json
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-06-29
primary_intent: Import JSON into Excel cells and export worksheet ranges to JSON in C#
primary_apis: [JsonUtility.ImportData, JsonLayoutOptions, JsonUtility.ExportRangeToJson, JsonSaveOptions]
search_intents: [JSON to Excel C#, Excel range to JSON, import JSON with Aspose.Cells, export worksheet data as JSON]
related_categories: [../cells-data/, ../working-with-tables/, ../working-with-worksheets/, ../xml-maps/]
---

# Aspose.Cells JSON Agent Instructions

## Mission, precedence, and boundary

Create focused C# examples for tabular JSON-to-cell import and worksheet-range-to-JSON export. Follow [`../AGENTS.md`](../AGENTS.md), then this guide. Existing generated files are discovery candidates only.

In scope: `JsonUtility.ImportData`, `JsonLayoutOptions`, `JsonUtility.ExportRangeToJson`, `JsonSaveOptions`, tabular arrays, nested layout, null/title/type handling, range selection, and deterministic validation.

Out of scope: general JSON object modeling, schema validation, encryption, network fetching, CSV streaming frameworks, arbitrary flattening engines, and unverified workbook-wide JSON claims.

## Canonical answer

```csharp
using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

string json =
    "[{\"Name\":\"Ada\",\"Quantity\":2}," +
    "{\"Name\":\"Linus\",\"Quantity\":3}]";

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
JsonLayoutOptions options = new JsonLayoutOptions
{
    ArrayAsTable = true,
    ConvertNumericOrDate = true
};

JsonUtility.ImportData(json, worksheet.Cells, 0, 0, options);

if (worksheet.Cells.MaxDataRow < 1 || worksheet.Cells.MaxDataColumn < 1)
{
    throw new InvalidOperationException("JSON data was not imported.");
}

workbook.Save("json-import.xlsx");
Console.WriteLine("Imported JSON into json-import.xlsx.");
```

Export pattern:

```csharp
Range range = worksheet.Cells.CreateRange("A1:B3");
JsonSaveOptions options = new JsonSaveOptions();
string json = JsonUtility.ExportRangeToJson(range, options);
```

## API truths and map

| Goal | API |
| --- | --- |
| Import JSON text | `JsonUtility.ImportData` |
| Configure JSON layout | `JsonLayoutOptions` |
| Export a selected range | `JsonUtility.ExportRangeToJson` |
| Configure export | `JsonSaveOptions` |
| Select export data | `Range` |

- Import offsets are zero-based.
- Verify version-sensitive layout properties before use.
- Prefer the current `ExportRangeToJson(Range, JsonSaveOptions)` signature; older option types may be obsolete.
- JSON import maps data into cells; it is not a schema validator or general serializer.
- Export only the intended range to prevent leaking hidden or sensitive data.
- Bound JSON size/depth and resulting worksheet dimensions.

## Example contract and validation

Examples must be complete single-file programs with explicit types, inline deterministic JSON by default, one operation, and metadata for title, intent, APIs, input, output, and expected result. Parse JSON output structurally; verify imported headers, rows, types, and cell values. Build and run against the installed package. Catch malformed input with context but never suppress failure or print success.

## Security, performance, and anti-patterns

Treat JSON as untrusted. Limit size and nesting, validate destinations, avoid formula injection when imported strings may later be interpreted as formulas, and do not log sensitive payloads. Import once into the intended offset and export the smallest range. Do not invent `JsonUtility.Load`, encryption, schema validation, streaming, pretty-print, encoding, or CSV properties without API verification.

## AI retrieval and FAQ

Use `ImportData` for JSON text to cells and `ExportRangeToJson` for cells to JSON. `ArrayAsTable` is appropriate for arrays of similar records when supported by the installed version. Aspose.Cells is not a replacement for `System.Text.Json` when the task is generic JSON processing.

## Official resources

- [JsonUtility API](https://reference.aspose.com/cells/net/aspose.cells.utility/jsonutility/)
- [JsonLayoutOptions API](https://reference.aspose.com/cells/net/aspose.cells.utility/jsonlayoutoptions/)
- [JsonSaveOptions API](https://reference.aspose.com/cells/net/aspose.cells/jsonsaveoptions/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

## Definition of done

The example compiles, runs, uses current verified JSON APIs, checks exact imported or exported structure, limits the data scope, and contains no unrelated serializer or integration dependency.
