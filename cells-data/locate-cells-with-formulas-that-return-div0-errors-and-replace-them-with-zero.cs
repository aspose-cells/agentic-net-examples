// Title: Replace #DIV/0! errors with zero in all worksheets of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an XLSX file with Aspose.Cells, forces formula calculation, scans every worksheet for cells containing the #DIV/0! error, and writes 0 into those cells. | Show a method using Aspose.Cells for .NET that iterates the used range of each sheet, detects error values, and replaces division‑by‑zero errors with the numeric value zero before saving the workbook. | Create a console‑application example that opens a workbook, evaluates all formulas, finds cells where Cell.IsErrorValue is true and Cell.StringValue contains "#DIV/0!", and sets the cell value to 0.
// Common Searches: asp.net aspose.cells replace division by zero error with zero in excel file | c# find #DIV/0! cells in workbook using Aspose.Cells | how to handle #DIV/0! errors when calculating formulas with Aspose.Cells | iterate used range and set zero for error cells Aspose.Cells C# | Aspose.Cells replace error values in Excel worksheets programmatically
// Tags: Aspose.Cells error value replacement | calculate formulas before error handling Aspose.Cells | iterate used range of worksheets C# | detect #DIV/0! cells Aspose.Cells | set numeric zero for Excel error cells | C# workbook error processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an input XLSX file, forces formula calculation, iterates each worksheet's used range, checks every cell for the '#DIV/0!' error, replaces those error cells with the numeric value 0, and saves the modified workbook to a new file.
    public class ReplaceDivZeroErrors
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure all formulas are calculated so that error values are materialized
                workbook.CalculateFormula();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Determine the used range of the worksheet
                    int maxRow = cells.MaxRow;
                    int maxColumn = cells.MaxColumn;

                    // Scan every cell within the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            Cell cell = cells[row, col];

                            // Check if the cell contains an error value
                            if (cell.IsErrorValue)
                            {
                                // Identify the specific #DIV/0! error by its string representation
                                if (!string.IsNullOrEmpty(cell.StringValue) && cell.StringValue.Contains("#DIV/0!"))
                                {
                                    // Replace the error with zero
                                    cell.PutValue(0);
                                }
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
