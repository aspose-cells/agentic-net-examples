// Title: How to unmerge a header row and add thin black borders to each cell in an exported Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an existing workbook, unmerges a specified header range (e.g., A1:D1), and applies thin black borders to every cell in that range for print‑ready output. | Create a reusable C# method using Aspose.Cells to remove merged header cells and set individual border styles before saving the workbook.
// Common Searches: aspnet unmerge merged header cells and set borders with Aspose.Cells | c# Aspose.Cells apply individual borders after unmerging header row | how to add thin black borders to each cell of a previously merged range in Excel using Aspose.Cells | printing Excel file with unmerged header and cell borders using Aspose.Cells .NET | remove merged cells from exported Excel and style header cells Aspose.Cells
// Tags: header range unmerge Aspose.Cells | thin black cell borders Aspose.Cells | cell style customization Excel .NET | merged cells removal Aspose.Cells | print formatting Excel workbook C#

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an exported Excel workbook, detects and removes a merged header range (A1:D1), then iterates over each cell in that range to apply thin black borders on all sides, and finally saves the modified workbook ready for printing.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "ExportedData.xlsx";
            const string outputPath = "ExportedData_WithHeaderBorders.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that was created during data export
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the header range that was previously merged (e.g., A1:D1)
            CellArea headerArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 0,        // Row 1
                EndColumn = 3      // Column D
            };

            // Unmerge the header range if it is merged
            // Aspose.Cells may not expose Cells.Unmerge in some versions; use MergedCells collection instead
            if (sheet.Cells.MergedCells.Contains(headerArea))
            {
                sheet.Cells.MergedCells.Remove(headerArea);
            }

            // Apply individual borders to each cell in the header range
            for (int row = headerArea.StartRow; row <= headerArea.EndRow; row++)
            {
                for (int col = headerArea.StartColumn; col <= headerArea.EndColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];
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
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
