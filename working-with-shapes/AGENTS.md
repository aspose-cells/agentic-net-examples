---
name: Aspose.Cells Shapes Agent
category: working-with-shapes
product: Aspose.Cells for .NET
language: C#
framework: .NET
repository: agentic-net-examples
parent: ../AGENTS.md
version: 3.0
last_reviewed: 2026-08-21
primary_intent: Create and manage Excel drawing shapes in C# without Microsoft Excel
primary_apis: [Worksheet.Shapes, ShapeCollection, Shape, TextBox, Picture, MsoDrawingType]
search_intents: [add Excel shape C#, create text box in Excel, format or position worksheet shape, add alternative text to Excel shape]
related_categories: [../working-with-images/, ../working-with-charts/, ../comments-and-notes/, ../working-with-worksheets/]
---

# Aspose.Cells Shapes Agent Instructions

## Mission and boundary

Create focused, runnable C# examples for worksheet drawing objects. Follow [`../AGENTS.md`](../AGENTS.md), then this file. Generated examples are candidates for validation, not authoritative API documentation.

In scope: AutoShapes, text boxes, WordArt, connectors, form controls, ActiveX/OLE objects, placement, grouping, z-order, resize, rotation, text, fill, line, links, locks, and alternative text. Pictures belong here only when common `Shape` behavior is dominant.

Out of scope: image rendering/compression, chart-specific behavior, comments, worksheet layout as the primary task, control activation, macros, UI automation, and external application integration.

## Canonical answer

```csharp
using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 240);
shape.Text = "Quarterly review";
shape.AlternativeText = "Rectangle labeled Quarterly review";

if (worksheet.Shapes.Count != 1 || shape.Text != "Quarterly review")
{
    throw new InvalidOperationException("The rectangle was not created.");
}

workbook.Save("excel-shape.xlsx");
Console.WriteLine("Created one labeled rectangle in excel-shape.xlsx.");
```

`AddRectangle` uses zero-based row, pixel top offset, zero-based column, pixel left offset, pixel height, and pixel width.

## API truths and map

| Goal | API |
| --- | --- |
| Access drawing objects | `Worksheet.Shapes` |
| Add common shapes | `ShapeCollection` add methods |
| Configure shared properties | `Shape` |
| Text box operations | `TextBox` / `Worksheet.TextBoxes` |
| Picture-specific operations | `Picture` / `Worksheet.Pictures` |
| Inspect drawing type | `MsoDrawingType` |
| Remove a shape | `ShapeCollection.RemoveAt` |

- Shapes are owned by a worksheet.
- Some add methods return an index; retrieve and cast the created object only through verified types.
- Anchors are row/column indexes while offsets and dimensions are pixels.
- Use specialized types when specialized behavior is required.
- `Shape.Text` is not the same as a linked-cell or formula-driven value.
- Verify collection state and round-trip serialization; rendering is supplemental validation.
- Never activate embedded controls, OLE objects, macros, or hyperlinks during validation.

## Contract, validation, and safety

Every example must use explicit types, one visible shape with positive dimensions, deterministic content/output, and metadata describing APIs and expected state. Build, run, assert collection count and semantic properties, save, and reopen when persistence matters.

Treat external images, hyperlinks, OLE payloads, and loaded workbooks as untrusted. Validate URI, host, timeout, content type, size, and paths. Cap shape counts and payload sizes. Do not claim alternative text alone provides full accessibility compliance.

## Anti-patterns and AI retrieval

Do not guess add overloads or enums, confuse anchors with offsets, depend on a fragile collection index after mutations, create zero-sized shapes, or catch an error and print success.

Use `Worksheet.Shapes` for generic shapes, `Worksheet.TextBoxes` for text boxes, and `Worksheet.Pictures` for picture-specific behavior. Use the images category when the goal is rendering rather than manipulating a drawing object.

## Official resources

- [Shapes documentation](https://docs.aspose.com/cells/net/managing-shapes/)
- [Shape API](https://reference.aspose.com/cells/net/aspose.cells.drawing/shape/)
- [ShapeCollection API](https://reference.aspose.com/cells/net/aspose.cells.drawing/shapecollection/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

## Definition of done

The example compiles, runs, creates or changes the claimed drawing object, verifies semantic state and persisted output, and performs no unsafe activation or unrelated integration.

