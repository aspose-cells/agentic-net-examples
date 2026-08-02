using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class UnmergeHeaderAndApplyBorders
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the exported data with a merged header.
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the merged header range (example: A1:D1).
            // Row and column indexes are zero‑based.
            int firstRow = 0;          // Row 1
            int firstColumn = 0;       // Column A
            int totalRows = 1;         // Only one row in the header
            int totalColumns = 4;      // Columns A to D

            // Unmerge the header range.
            cells.UnMerge(firstRow, firstColumn, totalRows, totalColumns);

            // Apply individual thin black borders to each cell that was part of the header.
            for (int row = firstRow; row < firstRow + totalRows; row++)
            {
                for (int col = firstColumn; col < firstColumn + totalColumns; col++)
                {
                    // Retrieve the current style of the cell.
                    Style style = cells[row, col].GetStyle();

                    // Set thin black borders on all four sides.
                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    style.Borders[BorderType.TopBorder].Color = Color.Black;
                    style.Borders[BorderType.BottomBorder].Color = Color.Black;
                    style.Borders[BorderType.LeftBorder].Color = Color.Black;
                    style.Borders[BorderType.RightBorder].Color = Color.Black;

                    // Apply the modified style back to the cell.
                    cells[row, col].SetStyle(style);
                }
            }

            // Save the workbook with the unmerged header and individual cell borders.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}