// Title: Replace #N/A and other Excel error values with a custom placeholder in a .NET workbook using Aspose.Cells
// Common Searches: Aspose.Cells C# replace #N/A error with custom text in Excel file | C# iterate through workbook cells and change error values to placeholder using Aspose.Cells | How to handle Excel error values like #N/A when saving with Aspose.Cells .NET | Replace all error cells in an XLSX using Aspose.Cells and save to a new file | Custom error handler for Excel formulas in Aspose.Cells C# example
// Tags: replace Excel error values Aspose.Cells C# | custom placeholder for #N/A cells Aspose.Cells | iterate workbook cells error handling Aspose.Cells | save modified workbook Aspose.Cells .NET | detect CellValueType.IsError Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsErrorHandler
{
    // The example loads an input XLSX workbook with Aspose.Cells, iterates over each worksheet's used range, checks each cell for the IsError type (including #N/A), replaces any error value with a user‑defined placeholder string, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string placeholder = "Data Not Available";
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Get the used range of the worksheet.
                    Aspose.Cells.Range usedRange = cells.MaxDisplayRange;
                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Loop through the used cells.
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Identify cells that contain an error value.
                            if (cell.Type == CellValueType.IsError)
                            {
                                // Replace any error (including #N/A) with the placeholder.
                                cell.PutValue(placeholder);
                            }
                        }
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
