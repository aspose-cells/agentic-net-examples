// Title: Merge cells F12:G12, apply a thick border style, and save as a macro‑enabled XLSM workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that merges the range F12:G12, creates a style with thick borders on all sides, applies the style to the merged range, and saves the workbook as an XLSM file. | Demonstrate how to format a merged cell block with uniform thick borders and export the result as a macro‑enabled spreadsheet using Aspose.Cells in a .NET application.
// Common Searches: how to merge a specific range and add thick borders using Aspose.Cells in C# | saving a workbook as macro enabled XLSM with Aspose.Cells .NET | apply border style to merged cells F12 G12 Aspose.Cells example | C# Aspose.Cells create styled macro enabled spreadsheet
// Tags: merge cells Aspose.Cells C# | thick border style Aspose.Cells | save workbook as XLSM Aspose.Cells | styled merged range Aspose.Cells .NET | macro enabled workbook generation Aspose.Cells

using System;
using Aspose.Cells;

// // This program creates a new workbook, merges cells F12:G12 on the first worksheet, applies a thick border style to the merged range, and saves the workbook as a macro‑enabled XLSM file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells F12:G12 (row index 11, column index 5)
            sheet.Cells.Merge(11, 5, 1, 2);

            // Create a style with thick borders on all sides
            Style style = workbook.CreateStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

            // Apply the style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("F12:G12");
            mergedRange.ApplyStyle(style, new StyleFlag() { All = true });

            // Save the workbook as a macro‑enabled XLSM file
            workbook.Save("output.xlsm", SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
