// Title: Read numeric cells from the first worksheet of an Excel file and calculate their average using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read all numeric cells from first worksheet and compute average | How to iterate over used range in an Excel file with Aspose.Cells and calculate mean value | Example of handling password‑protected Excel workbook while averaging numeric data using Aspose.Cells | Console application to calculate average of numbers in an exported Apple Numbers .xlsx using Aspose.Cells
// Tags: average numeric cells Aspose.Cells | iterate used range worksheet C# | extract numeric values Excel Aspose.Cells | handle password protected workbook exception Aspose.Cells | load workbook from file path Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The C# console program verifies the presence of an input.xlsx file, loads it with Aspose.Cells, gracefully handles missing or password‑protected workbooks, accesses the first worksheet, iterates through the used rows and columns to collect numeric cell values, computes their average, and writes the result to the console while managing any processing errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook (no password supplied)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            // Detect password‑protected workbook by inspecting the exception message
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Please provide a valid password.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
            return;
        }

        try
        {
            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Collect all numeric values from the used range
            List<double> numericValues = new List<double>();
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        numericValues.Add(cell.DoubleValue);
                    }
                }
            }

            // Calculate the average of the collected numbers
            double average = numericValues.Count > 0 ? numericValues.Average() : 0.0;

            // Output the result to the console
            Console.WriteLine($"Average of numeric cells: {average}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
