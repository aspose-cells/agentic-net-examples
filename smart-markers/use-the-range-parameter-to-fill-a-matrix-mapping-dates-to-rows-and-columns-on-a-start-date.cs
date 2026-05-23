using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDateMatrixDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Cells cells = workbook.Worksheets[0].Cells;

                // Define matrix dimensions (including header row and column)
                int totalRows = 6;    // 1 header row + 5 date rows
                int totalColumns = 6; // 1 header column + 5 date columns

                // Define the start date for the matrix
                DateTime startDate = new DateTime(2023, 1, 1);

                // Prepare a 2‑dimensional object array to hold the matrix data
                object[,] matrix = new object[totalRows, totalColumns];

                // Fill the top‑left cell (intersection of headers) with a label
                matrix[0, 0] = "Date\\Date";

                // Fill the first row with consecutive dates (column headers)
                for (int col = 1; col < totalColumns; col++)
                {
                    matrix[0, col] = startDate.AddDays(col - 1);
                }

                // Fill the first column with consecutive dates (row headers)
                for (int row = 1; row < totalRows; row++)
                {
                    matrix[row, 0] = startDate.AddDays(row - 1);
                }

                // Fill the inner matrix cells with a simple calculation (day offset sum)
                for (int row = 1; row < totalRows; row++)
                {
                    for (int col = 1; col < totalColumns; col++)
                    {
                        DateTime rowDate = (DateTime)matrix[row, 0];
                        DateTime colDate = (DateTime)matrix[0, col];
                        matrix[row, col] = (colDate - rowDate).Days;
                    }
                }

                // Create a range that covers the entire matrix area and assign the array
                AsposeRange range = cells.CreateRange(0, 0, totalRows, totalColumns);
                range.Value = matrix;

                // Apply a date format to the header rows/columns
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // Built‑in date format (e.g., mm/dd/yyyy)
                StyleFlag flag = new StyleFlag { NumberFormat = true };

                // Apply style to first row (excluding top‑left cell)
                cells.CreateRange(0, 1, 1, totalColumns - 1).ApplyStyle(dateStyle, flag);
                // Apply style to first column (excluding top‑left cell)
                cells.CreateRange(1, 0, totalRows - 1, 1).ApplyStyle(dateStyle, flag);

                // Define output file path
                string outputPath = "DateMatrix.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}