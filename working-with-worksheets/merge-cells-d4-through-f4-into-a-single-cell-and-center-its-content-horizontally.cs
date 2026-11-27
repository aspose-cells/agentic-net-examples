// Title: How to merge cells D4‑F4 and horizontally center the text using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge range D4:F4 and align text to center | C# code to merge cells D4 to F4 in Excel and set horizontal alignment with Aspose | center text in merged Excel cells using Aspose.Cells .NET library | merge cells D4-F4 in a worksheet and apply center alignment with Aspose.Cells | Aspose.Cells example for merging cells and setting style in C#
// Tags: Aspose.Cells merge cells programmatically | horizontal alignment for merged Excel cells C# | apply style to merged range Aspose.Cells | create .xlsx with merged cells using Aspose | Excel worksheet cell merging API .NET

using System;
using Aspose.Cells;

// The sample creates a new workbook, merges cells D4 through F4 on the first worksheet, inserts the text "Merged D4:F4", applies a style that centers the text horizontally (and vertically), and saves the file as MergedCells_D4_F4.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range D4:F4 (row index 3, column index 3, 1 row, 3 columns)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange(3, 3, 1, 3);

            // Merge the cells in the range into a single cell
            range.Merge();

            // Set a value in the merged cell (upper‑left cell of the range)
            worksheet.Cells[3, 3].PutValue("Merged D4:F4");

            // Center the content horizontally and vertically
            Style style = worksheet.Cells[3, 3].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            worksheet.Cells[3, 3].SetStyle(style);

            // Save the workbook
            string outputPath = "MergedCells_D4_F4.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
