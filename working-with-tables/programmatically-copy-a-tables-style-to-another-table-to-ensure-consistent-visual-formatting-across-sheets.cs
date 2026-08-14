// Title: Copy a table style between worksheets with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a source ListObject, apply a built‑in TableStyleMedium2, then assign the same TableStyleName to a table on another sheet, ensuring identical visual formatting before saving the workbook.
// Keywords: Aspose.Cells C# | copy Excel table style | ListObject TableStyleName | apply built‑in table style programmatically | TableStyleMedium2 Aspose | worksheet table formatting | Aspose.Cells example | Excel style cloning .NET
// Common Searches: Aspose.Cells copy table formatting to another sheet | C# assign same TableStyleName to multiple ListObjects | how to reuse Excel table style with Aspose.Cells | programmatically set built‑in table style in .NET | duplicate table appearance across worksheets Aspose
// Developer Intent: Reuse the visual formatting of one ListObject for another table on a different worksheet.
// Use Cases: Standardize branding by applying a single table style to all report sheets. | Generate new worksheets dynamically while preserving a predefined table appearance. | Clone formatting when creating month‑over‑month comparison tabs without manual styling.
// AI Prompts: Show C# code that copies a ListObject's TableStyleName to another table using Aspose.Cells. | Explain how to retrieve all built‑in table styles in Aspose.Cells and apply a chosen style to several tables. | Provide an example of cloning table formatting across multiple worksheets in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a source ListObject, apply a built‑in TableStyleMedium2, then assign the same TableStyleName to a table on another sheet, ensuring identical visual formatting before saving the workbook.
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

        // Add a table on the source sheet
        int srcTableIdx = srcSheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject srcTable = srcSheet.ListObjects[srcTableIdx];
        srcTable.ShowTableStyleFirstColumn = true;
        srcTable.ShowTableStyleLastColumn = true;

        // Apply a built‑in table style to the source table
        TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
        TableStyle builtinStyle = tableStyles.GetBuiltinTableStyle(TableStyleType.TableStyleMedium2);
        srcTable.TableStyleName = builtinStyle.Name;

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

        // Add a table on the destination sheet
        int destTableIdx = destSheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject destTable = destSheet.ListObjects[destTableIdx];
        destTable.ShowTableStyleFirstColumn = true;
        destTable.ShowTableStyleLastColumn = true;

        // Copy the style from the source table to the destination table
        destTable.TableStyleName = srcTable.TableStyleName;

        // Save the workbook
        workbook.Save("TableStyleCopied.xlsx");
    }
}
