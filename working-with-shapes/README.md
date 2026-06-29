---
title: Create and Manage Excel Shapes in C# with Aspose.Cells
description: Add, format, position, link, group, and remove Excel shapes, text boxes, controls, WordArt, and pictures in C#.
product: Aspose.Cells for .NET
category: working-with-shapes
language: C#
last_reviewed: 2026-06-29
---

# Create and Manage Excel Shapes in C# with Aspose.Cells

Use `Worksheet.Shapes` and specialized drawing types to create and manage Excel shapes without Microsoft Excel.

| Repository fact | Value |
| --- | --- |
| Examples | 532 standalone `.cs` files |
| Primary APIs | `Worksheet.Shapes`, `ShapeCollection`, `Shape`, `TextBox`, `Picture` |
| Excel required | No |
| Agent guidance | [`agents.md`](agents.md) |
| Catalog | [`../index.json`](../index.json) |

## Quick answer: How do I add a shape to Excel in C#?

```csharp
using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 240);
shape.Text = "Quarterly review";
shape.AlternativeText = "Rectangle labeled Quarterly review";

if (worksheet.Shapes.Count != 1)
{
    throw new InvalidOperationException("Shape creation failed.");
}

workbook.Save("excel-shape.xlsx");
Console.WriteLine("Created one labeled rectangle in excel-shape.xlsx.");
```

## API choice

| Need | API |
| --- | --- |
| Generic shapes and ordering | `Worksheet.Shapes` |
| Shared text/fill/line/placement | `Shape` |
| Text boxes | `Worksheet.TextBoxes`, `TextBox` |
| Pictures | `Worksheet.Pictures`, `Picture` |
| Drawing type | `MsoDrawingType` |

## Featured examples

- [Add a worksheet text box](add-a-new-textbox-to-the-worksheet-at-the-specified-cell-coordinates.cs)
- [Set alternative text on a shape](add-a-shape-and-set-alternative-text-describing-its-purpose-for-accessibility-compliance.cs)
- [Add a checkbox linked to cell B2](add-a-checkbox-control-to-a-worksheet-and-link-its-state-to-cell-b2.cs)
- [Rotate a picture 90 degrees](add-a-picture-and-rotate-it-ninety-degrees-clockwise-to-align-with-column-orientation.cs)
- [Adjust shape z-order](adjust-a-shapes-z-order-to-just-above-a-specific-existing-shape.cs)
- [Add a hyperlink to a text box](add-a-hyperlink-to-the-textbox-that-opens-a-web-page-when-the-shape-is-clicked.cs)

Specialized examples may be version-sensitive. Validate APIs and never activate controls, links, OLE objects, or macros during automated checks.

## Getting started

```bash
dotnet new console -n ExcelShapeExample
cd ExcelShapeExample
dotnet add package Aspose.Cells
dotnet build
dotnet run
```

Shape add methods combine zero-based cell anchors with pixel offsets and dimensions. Verify positive dimensions, collection count, text, alternative text, placement, and round-trip persistence.

## FAQ

**Are shapes stored at workbook level?** No. They belong to a worksheet.

**When should I use `Picture` instead of `Shape`?** Use `Picture` for picture-specific properties and `Shape` for shared drawing behavior.

**Does alternative text guarantee accessibility compliance?** No. It is one useful property within a broader accessibility review.

**Can I validate a shape only by rendering it?** Prefer semantic assertions first; rendering can supplement them.

## Related and official resources

- [`working-with-images`](../working-with-images/)
- [`working-with-charts`](../working-with-charts/)
- [`comments-and-notes`](../comments-and-notes/)
- [Managing Shapes](https://docs.aspose.com/cells/net/managing-shapes/)
- [Shape API](https://reference.aspose.com/cells/net/aspose.cells.drawing/shape/)
- [ShapeCollection](https://reference.aspose.com/cells/net/aspose.cells.drawing/shapecollection/)
- [Aspose.Cells NuGet](https://www.nuget.org/packages/Aspose.Cells/)

Review the repository [`LICENSE`](../LICENSE) and Aspose licensing terms before production use.
