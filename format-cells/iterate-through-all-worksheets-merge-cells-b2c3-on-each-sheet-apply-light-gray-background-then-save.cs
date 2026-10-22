// Title: Merge B2:C3 cells on every worksheet and apply a light‑gray fill with Aspose.Cells for .NET
// Common Searches: C# merge B2:C3 range on all worksheets using Aspose.Cells | apply solid light gray background to a merged range with Aspose.Cells .NET | loop through workbook sheets to format cells in Aspose.Cells | how to set fill color for merged cells in Aspose.Cells C# | Aspose.Cells example for merging cells and styling across multiple sheets
// Tags: merge cells B2:C3 Aspose.Cells | gray fill merged range Aspose.Cells | format range B2:C3 Aspose.Cells | loop through worksheets Aspose.Cells | create workbook multiple sheets Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

// The code creates a new workbook, adds three worksheets, iterates over each sheet, merges the range B2:C3, defines a style with a light‑gray solid fill, applies the style to the merged range, and saves the workbook as MergedSheets.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add worksheets (optional)
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Merge cells B2:C3 (zero‑based indexes: row 1, column 1 to row 2, column 2)
                sheet.Cells.Merge(1, 1, 2, 2);

                // Create a style with light gray background
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightGray;
                style.Pattern = BackgroundType.Solid;

                // Apply the style to the merged range
                Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("B2:C3");
                mergedRange.ApplyStyle(style, new StyleFlag { All = true });
            }

            // Save the workbook to a file
            workbook.Save("MergedSheets.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
