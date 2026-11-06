// Title: Create a macro‑enabled XLSM workbook, merge cells F12:G12, and apply a thick black border using Aspose.Cells for .NET
// Common Searches: how to merge a range of cells and add a thick border with Aspose.Cells in C# | saving a workbook as a macro‑enabled XLSM file using Aspose.Cells .NET | apply a custom style with thick borders to a merged range in Aspose.Cells | Aspose.Cells C# example for merging cells F12 to G12 and styling borders | create and style merged cells then export as .xlsm with Aspose.Cells
// Tags: merge cells with Aspose.Cells C# | apply thick border style to merged range | save workbook as macro enabled XLSM | create custom cell style using CellBorderType.Thick | style merged cells using StyleFlag All

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a new workbook, merges cells F12:G12, applies a thick black border style to the merged range, and saves the file as a macro‑enabled XLSM.
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

            // Merge cells F12:G12 (row 11, columns 5 and 6; zero‑based indexing)
            sheet.Cells.Merge(11, 5, 1, 2);

            // Create a style with a thick border on all sides
            Style thickBorderStyle = workbook.CreateStyle();
            thickBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

            // Optional: set border color
            thickBorderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
            thickBorderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
            thickBorderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
            thickBorderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

            // Apply the style to the merged range
            AsposeRange mergedRange = sheet.Cells.CreateRange("F12:G12");
            mergedRange.ApplyStyle(thickBorderStyle, new StyleFlag() { All = true });

            // Define output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.xlsm");

            // Save the workbook as a macro‑enabled XLSM file
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
