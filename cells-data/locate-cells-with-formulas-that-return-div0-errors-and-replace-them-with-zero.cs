// Title: C# – Replace #DIV/0! Errors with Zero Using Aspose.Cells
// Description: Loads a workbook, forces formula calculation, scans the used range of the first worksheet, detects cells containing the #DIV/0! error, substitutes each error with 0, and saves the updated file.
// Keywords: Aspose.Cells | C# | #DIV/0! error | replace Excel error with zero | formula recalculation | iterate used cells | Excel error handling | bulk replace errors | Excel automation
// Common Searches: Aspose.Cells replace #DIV/0! with 0 | C# find and fix division by zero errors in Excel | How to change Excel error values using Aspose.Cells | Iterate over used range Aspose.Cells C# | Calculate formulas before fixing errors Aspose
// Developer Intent: Replace every #DIV/0! error in the workbook with the numeric value 0.
// Use Cases: Prepare financial reports by removing division‑by‑zero errors that break downstream calculations. | Clean data before importing Excel files into systems that reject error values. | Run batch processing on multiple workbooks to ensure they contain only numeric values.
// AI Prompts: Generate C# code with Aspose.Cells that finds all #DIV/0! cells and replaces them with zero. | Explain how IsErrorValue and StringValue can be combined to identify specific Excel error types in Aspose.Cells. | Show an alternative approach using Cells.Find to locate and replace #DIV/0! errors in a worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, forces formula calculation, scans the used range of the first worksheet, detects cells containing the #DIV/0! error, substitutes each error with 0, and saves the updated file.
    public class ReplaceDivZeroErrors
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure all formulas are evaluated so error values are up‑to‑date
                workbook.CalculateFormula();

                // Access the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Iterate through all used cells in the worksheet
                for (int row = 0; row <= cells.MaxRow; row++)
                {
                    for (int col = 0; col <= cells.MaxColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell contains a #DIV/0! error
                        if (cell.IsErrorValue && cell.StringValue == "#DIV/0!")
                        {
                            // Replace the error with zero
                            cell.PutValue(0);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
