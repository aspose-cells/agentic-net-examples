// Title: Calculate a 10‑row sliding moving average for every column using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells code to add a rolling 10‑row average formula to each column in an existing workbook | how to programmatically insert AVERAGE formulas with a dynamic range for every column using Aspose.Cells | create a new worksheet that contains sliding window averages for all columns with Aspose.Cells .NET | load an Excel file, copy headers, and apply a ten‑row moving average across columns in C# | Aspose.Cells example for generating rolling averages per column and saving to a separate sheet
// Tags: Aspose.Cells sliding window moving average | C# set dynamic AVERAGE formula per column | generate rolling averages worksheet Aspose.Cells | copy header row and apply formulas .NET | calculate ten‑row rolling average Excel C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, adds a new sheet named "MovingAverage", copies the original header row, and inserts an AVERAGE formula in each column that computes a ten‑row sliding window starting at row 10. The workbook is then saved with the calculated averages.
class MovingAverageCalculator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            Workbook workbook;

            // Load the workbook; handle password‑protected files gracefully
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex) when (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is password protected and cannot be opened without a password.");
                return;
            }

            // Get the first worksheet containing the data
            Worksheet dataSheet = workbook.Worksheets[0];

            // Determine the used range
            int lastRow = dataSheet.Cells.MaxDataRow;          // zero‑based index of the last row with data
            int lastColumn = dataSheet.Cells.MaxDataColumn;    // zero‑based index of the last column with data

            // Add a new worksheet to store moving‑average results
            Worksheet maSheet = workbook.Worksheets.Add("MovingAverage");

            // Loop through each column in the source sheet
            for (int col = 0; col <= lastColumn; col++)
            {
                // Copy the header (assumed to be in the first row, index 0)
                maSheet.Cells[0, col].PutValue(dataSheet.Cells[0, col].StringValue);

                // Get the column letter for building Excel formulas (e.g., A, B, …)
                string colLetter = CellsHelper.ColumnIndexToName(col);

                // Apply moving‑average formula starting from the 10th data row (Excel row 10, zero‑based index 9)
                for (int row = 9; row <= lastRow; row++)
                {
                    int excelRow = row + 1;          // Excel rows are 1‑based
                    int startRow = excelRow - 9;     // start of the 10‑row window (inclusive)

                    // Build the AVERAGE formula, e.g., =AVERAGE(A2:A11)
                    string formula = $"AVERAGE({colLetter}{startRow}:{colLetter}{excelRow})";

                    // Set the formula in the moving‑average sheet
                    maSheet.Cells[row, col].Formula = formula;
                }
            }

            // Save the workbook with the new sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
