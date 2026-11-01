// Title: Unmerge cells B1:B4 and apply thin black borders using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unmerge cells in column B and set individual borders | C# code to split merged range B1:B4 and apply thin black borders with Aspose.Cells | How to restore cell borders after unmerging cells using Aspose.Cells for .NET | Unmerge specific range B1:B4 in an Excel workbook with Aspose.Cells and format borders | Aspose.Cells example for unmerging cells and applying border style in C#
// Tags: Aspose.Cells unmerge specific range | Aspose.Cells set thin borders | Excel unmerge cells with Aspose.Cells | C# apply cell borders after unmerge | Aspose.Cells range unmerge example

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The C# program loads an existing Excel workbook, checks for the file, defines the B1:B4 range, unmerges any merged cells within that range, applies thin black borders to each individual cell, and saves the modified workbook while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range B1:B4 (zero‑based indices)
            int startRow = 0;      // B1
            int endRow = 3;        // B4
            int column = 1;        // Column B

            // Unmerge the cells in the specified range if they are merged
            for (int row = startRow; row <= endRow; row++)
            {
                Cell cell = sheet.Cells[row, column];
                if (cell.IsMerged)
                {
                    // Unmerge the merged range that this cell belongs to
                    AsposeRange mergedRange = cell.GetMergedRange();
                    mergedRange.UnMerge();
                }
            }

            // Apply thin black borders to each individual cell in the range
            for (int row = startRow; row <= endRow; row++)
            {
                Cell cell = sheet.Cells[row, column];
                Style style = cell.GetStyle();

                // Top border
                style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.TopBorder].Color = Color.Black;

                // Bottom border
                style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.BottomBorder].Color = Color.Black;

                // Left border
                style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.LeftBorder].Color = Color.Black;

                // Right border
                style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.RightBorder].Color = Color.Black;

                cell.SetStyle(style);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
