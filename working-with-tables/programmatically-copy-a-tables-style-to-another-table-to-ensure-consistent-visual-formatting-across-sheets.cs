// Title: Copy a Table Style from One Worksheet to Another with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a source ListObject with a built‑in table style, add a destination ListObject on a different sheet, and transfer the TableStyleName, first‑column, and last‑column highlight settings programmatically before saving the workbook.
// Keywords: Aspose.Cells copy table style C# | transfer ListObject formatting .NET | apply same Excel table style multiple sheets | Aspose.Cells TableStyleName example | C# copy table visual formatting
// Common Searches: how to copy Aspose.Cells table style between worksheets | C# copy ListObject style to another sheet | duplicate Excel table formatting with Aspose.Cells | apply TableStyleMedium9 to multiple tables programmatically | Aspose.Cells copy table appearance code
// Developer Intent: Replicate the visual formatting of a source table on a different worksheet using Aspose.Cells for .NET.
// Use Cases: Generate multi‑sheet reports where every table follows the corporate table style for a consistent look. | Create a template workbook and programmatically add new tables that automatically inherit the template’s style and column highlights. | Synchronize table appearance after cloning or moving tables between worksheets in an automated data‑processing workflow.
// AI Prompts: Show C# code that copies all table style properties, including banded rows and columns, from one Aspose.Cells ListObject to another. | Explain how to transfer a table’s style together with its conditional formatting rules between worksheets using Aspose.Cells for .NET. | Provide a reusable method that accepts a source ListObject and a destination ListObject and applies the source’s TableStyleName, first‑column, and last‑column settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a source ListObject with a built‑in table style, add a destination ListObject on a different sheet, and transfer the TableStyleName, first‑column, and last‑column highlight settings programmatically before saving the workbook.
class CopyTableStyleDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ---------- Source sheet ----------
        Worksheet srcSheet = workbook.Worksheets[0];
        srcSheet.Name = "Source";

        // Fill source data
        srcSheet.Cells["A1"].PutValue("Product");
        srcSheet.Cells["B1"].PutValue("Price");
        srcSheet.Cells["A2"].PutValue("Apple");
        srcSheet.Cells["B2"].PutValue(1.2);
        srcSheet.Cells["A3"].PutValue("Banana");
        srcSheet.Cells["B3"].PutValue(0.8);

        // Add a table to the source sheet
        int srcTableIdx = srcSheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject srcTable = srcSheet.ListObjects[srcTableIdx];

        // Apply a built‑in table style to the source table
        srcTable.TableStyleName = "TableStyleMedium9";
        srcTable.ShowTableStyleFirstColumn = true;
        srcTable.ShowTableStyleLastColumn = true;

        // ---------- Destination sheet ----------
        Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        destSheet.Name = "Destination";

        // Fill destination data (same structure)
        destSheet.Cells["A1"].PutValue("Product");
        destSheet.Cells["B1"].PutValue("Price");
        destSheet.Cells["A2"].PutValue("Orange");
        destSheet.Cells["B2"].PutValue(1.5);
        destSheet.Cells["A3"].PutValue("Grape");
        destSheet.Cells["B3"].PutValue(2.0);

        // Add a table to the destination sheet
        int destTableIdx = destSheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject destTable = destSheet.ListObjects[destTableIdx];

        // Copy the style settings from the source table to the destination table
        destTable.TableStyleName = srcTable.TableStyleName;
        destTable.ShowTableStyleFirstColumn = srcTable.ShowTableStyleFirstColumn;
        destTable.ShowTableStyleLastColumn = srcTable.ShowTableStyleLastColumn;

        // Save the workbook
        workbook.Save("TableStyleCopied.xlsx");
    }
}
