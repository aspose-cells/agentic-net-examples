// Title: Aspose.Cells for .NET – Replace #DIV/0! Errors with Zero in an Excel Workbook
// Description: Loads a workbook, forces formula recalculation, scans every used cell across all worksheets, detects the #DIV/0! error, substitutes it with the numeric value 0, and saves the cleaned file.
// Keywords: Aspose.Cells replace DIV/0 error | C# set #DIV/0! to zero | Excel error handling .NET | cell.IsErrorValue Aspose example | remove division by zero Aspose.Cells | Aspose.Cells formula error cleanup | global Excel data sanitization
// Common Searches: how to change #DIV/0! to 0 using Aspose.Cells | Aspose.Cells replace division by zero error | C# iterate worksheets and fix Excel errors | Aspose.Cells calculate formulas then clean errors | replace Excel error values programmatically .NET
// Developer Intent: Find and replace cells that contain the #DIV/0! error with a numeric zero.
// Use Cases: Clean financial reports before sharing so division‑by‑zero cells display 0 instead of an error. | Prepare Excel uploads for ERP systems that reject error values. | Automate batch processing of multiple workbooks to ensure all #DIV/0! cells are numeric.
// AI Prompts: Show how to replace any Excel error (e.g., #N/A, #VALUE!) with a configurable default using Aspose.Cells. | Add logging that records the address of each cell changed from #DIV/0! to zero. | Explain how to skip the CalculateFormula step when the workbook is already up‑to‑date while still fixing #DIV/0! errors.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, forces formula recalculation, scans every used cell across all worksheets, detects the #DIV/0! error, substitutes it with the numeric value 0, and saves the cleaned file.
    public class ReplaceDivZeroWithZero
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Ensure all formulas are calculated so error values are up‑to‑date
            workbook.CalculateFormula();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Scan the used range of the worksheet
                int maxRow = cells.MaxRow;
                int maxCol = cells.MaxColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Replace #DIV/0! errors with numeric zero
                        if (cell.IsErrorValue && cell.StringValue == "#DIV/0!")
                        {
                            cell.PutValue(0);
                        }
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
