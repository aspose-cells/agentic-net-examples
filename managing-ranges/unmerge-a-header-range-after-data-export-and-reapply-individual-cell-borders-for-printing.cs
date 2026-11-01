// Title: How to unmerge a header row and reapply individual thin borders to each cell in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unmerge merged header cells and add thin borders for printing | C# code to split merged cells A1:D1 and set individual borders with Aspose.Cells | Unmerge range and apply cell borders after exporting Excel using Aspose.Cells .NET | How to programmatically remove a merged header and set borders in an Aspose.Cells workbook | Aspose.Cells example for unmerging a range and styling borders in C#
// Tags: unmerge header range Aspose.Cells C# | set individual cell borders Aspose.Cells | range.UnMerge usage example | cell style border configuration .NET | address to row column conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHeaderUnmerge
{
    // The sample loads an existing workbook, unmerges a specified header range (e.g., A1:D1), reapplies thin borders to each cell in that range, and saves the modified file ready for printing.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains the exported data
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the header range that was previously merged (e.g., A1:D1)
                string startAddress = "A1";
                string endAddress = "D1";

                // Convert the addresses to zero‑based row and column indices
                (int startRow, int startColumn) = GetRowColumnFromAddress(startAddress);
                (int endRow, int endColumn) = GetRowColumnFromAddress(endAddress);

                // Calculate the size of the range for unmerge
                int totalRows = endRow - startRow + 1;
                int totalColumns = endColumn - startColumn + 1;

                // Unmerge the header range using a Range object
                Aspose.Cells.Range range = worksheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
                range.UnMerge();

                // Reapply individual borders to each cell in the former merged area
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        // Retrieve the current style of the cell
                        Style cellStyle = worksheet.Cells[row, col].GetStyle();

                        // Set thin borders on all sides
                        cellStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                        cellStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                        cellStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                        cellStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                        // Apply the modified style back to the cell
                        worksheet.Cells[row, col].SetStyle(cellStyle);
                    }
                }

                // Save the workbook ready for printing
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static (int row, int column) GetRowColumnFromAddress(string address)
        {
            // Separate letters (column) from numbers (row)
            int i = 0;
            while (i < address.Length && char.IsLetter(address[i]))
                i++;

            string columnPart = address.Substring(0, i).ToUpperInvariant();
            string rowPart = address.Substring(i);

            // Convert column letters to zero‑based index
            int columnIndex = CellsHelper.ColumnNameToIndex(columnPart);

            // Convert row number to zero‑based index
            if (!int.TryParse(rowPart, out int rowNumber) || rowNumber < 1)
                throw new ArgumentException($"Invalid cell address: {address}");

            int rowIndex = rowNumber - 1;

            return (rowIndex, columnIndex);
        }
    }
}
