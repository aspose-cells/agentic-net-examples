// Title: Apply Built‑In Table Style Matching Workbook Theme to a ListObject (Aspose.Cells for .NET)
// Description: Demonstrates how to create a worksheet, add a ListObject, retrieve a built‑in TableStyleMedium2 that follows the workbook's current theme via TableStyleCollection, assign the style, enable first/last column highlights and row stripes, and save the file using Aspose.Cells for C#.
// Keywords: Aspose.Cells C# | apply built‑in table style | TableStyleMedium2 | TableStyleCollection | ListObject styling | Excel theme‑aware style | row stripe formatting | first column highlight | last column highlight | save workbook Aspose
// Common Searches: Aspose.Cells apply theme table style C# | How to set ListObject TableStyleName to built‑in style | Retrieve TableStyleMedium2 with Aspose.Cells | Enable row stripes and column highlights in Aspose.Cells table | C# code for applying built‑in Excel table style
// Developer Intent: Use Aspose.Cells to apply a theme‑compatible built‑in table style to a newly added ListObject.
// Use Cases: Standardize the appearance of inventory tables across a workbook by applying a built‑in medium style that respects the workbook theme. | Improve readability of data tables by turning on first/last column shading and alternating row colors. | Create a reusable routine that selects any TableStyleType, applies it to a table, and configures visual options for multiple worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that adds a ListObject to a worksheet, selects a TableStyleLight1 built‑in style, applies it, and enables column banding. | Write a method for Aspose.Cells that accepts a worksheet, a range address, and a TableStyleType, then creates a table, applies the corresponding built‑in style, and sets row stripe and column highlight options.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a worksheet, add a ListObject, retrieve a built‑in TableStyleMedium2 that follows the workbook's current theme via TableStyleCollection, assign the style, enable first/last column highlights and row stripes, and save the file using Aspose.Cells for C#.
class ApplyBuiltinTableStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(20);

        // Add a table (ListObject) that covers the data range
        int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Retrieve a built‑in table style that matches the workbook's theme
        TableStyleCollection styleCollection = workbook.Worksheets.TableStyles;
        TableStyle builtinStyle = styleCollection.GetBuiltinTableStyle(TableStyleType.TableStyleMedium2);

        // Apply the built‑in style to the table
        table.TableStyleName = builtinStyle.Name;

        // Optional visual enhancements
        table.ShowTableStyleFirstColumn = true;
        table.ShowTableStyleLastColumn = true;
        table.ShowTableStyleRowStripes = true;

        // Save the workbook
        workbook.Save("AppliedBuiltinTableStyle.xlsx");
    }
}
