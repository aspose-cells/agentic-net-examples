// Title: Unmerge Header Row and Add Thin Black Borders to Each Cell with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, unmerges a header range (e.g., A1:D1), applies thin black borders to every cell that was part of the merged area, and saves the file ready for printing.
// Keywords: Aspose.Cells | C# | unmerge range | cell borders | thin black border | Excel header styling | print‑ready workbook | range.UnMerge | style borders | .NET Excel automation
// Common Searches: Aspose.Cells unmerge header row C# | How to add borders after unmerging cells in Aspose.Cells | C# Aspose.Cells set thin borders for each cell | Unmerge merged cells and apply borders Aspose.Cells .NET | Print‑ready Excel with individual cell borders using Aspose
// Developer Intent: The developer wants to programmatically break a merged header into separate cells and give each resulting cell its own thin black border before the workbook is printed or exported.
// Use Cases: Create print‑ready Excel reports where merged titles must be split and bordered. | Prepare data exports for PDF conversion with consistent cell outlines. | Standardize worksheet appearance across generated files by automating unmerge and border styling.
// AI Prompts: Write C# code using Aspose.Cells to unmerge a specified range and apply thin black borders to each resulting cell. | Show how to preserve existing cell formatting while adding borders after unmerging a header row in Aspose.Cells. | Explain step‑by‑step how to iterate over a range and set BorderType.Thin for all sides in Aspose.Cells .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, unmerges a header range (e.g., A1:D1), applies thin black borders to every cell that was part of the merged area, and saves the file ready for printing.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains exported data
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Address of the merged header range (e.g., A1:D1)
            string headerRangeAddress = "A1:D1";

            // Create the range using Aspose.Cells.Range to avoid ambiguity with System.Range
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(headerRangeAddress);
            headerRange.UnMerge();

            // Apply individual borders to each cell that was part of the merged header
            int firstRow = headerRange.FirstRow;
            int firstColumn = headerRange.FirstColumn;
            int totalRows = headerRange.RowCount;
            int totalColumns = headerRange.ColumnCount;

            for (int r = firstRow; r < firstRow + totalRows; r++)
            {
                for (int c = firstColumn; c < firstColumn + totalColumns; c++)
                {
                    Cell cell = sheet.Cells[r, c];
                    Style style = cell.GetStyle();

                    // Set thin black borders on all sides
                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    style.Borders[BorderType.TopBorder].Color = Color.Black;
                    style.Borders[BorderType.BottomBorder].Color = Color.Black;
                    style.Borders[BorderType.LeftBorder].Color = Color.Black;
                    style.Borders[BorderType.RightBorder].Color = Color.Black;

                    cell.SetStyle(style);
                }
            }

            // Save the workbook ready for printing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
