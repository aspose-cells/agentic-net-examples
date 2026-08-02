// Title: Preserve Excel Formulas While Merging Cells – Set Automatic Calculation Mode with Aspose.Cells for .NET
// Description: Shows how to load (or create) a workbook, force FormulaSettings.CalculationMode to Automatic before and after merging the range A1:B2 on the first worksheet, and save the file so that all existing formulas stay intact.
// Keywords: Aspose.Cells | C# | merge cells | preserve formulas | automatic calculation mode | Workbook merge | Excel automation | FormulaSettings | CalcModeType.Automatic
// Common Searches: Aspose.Cells keep formulas after merging cells | set calculation mode automatic before merge Aspose.Cells | C# merge Excel range without breaking formulas | preserve Excel formulas Aspose.Cells .NET | automatic calculation mode merge cells example
// Developer Intent: Maintain formula integrity by keeping the workbook’s calculation mode set to Automatic before and after a cell‑merge operation.
// Use Cases: Load an existing workbook, enforce Automatic calculation, merge A1:B2, and save the result. | Create a new workbook when the source file is missing, apply the same Automatic setting, merge cells, and preserve any default formulas. | Process multiple worksheets in a loop, setting Automatic mode before each merge to ensure formulas remain correct across the entire file.
// AI Prompts: Write C# code using Aspose.Cells that merges cells A1:B2 while preserving all formulas by setting the workbook’s calculation mode to Automatic before and after the merge. | Generate a method that checks for an input Excel file, creates a fallback workbook if missing, forces Automatic calculation mode, merges a specified range, and saves the workbook. | Provide an example that iterates through every worksheet in a workbook, merges a given cell range on each sheet, and keeps formulas intact by using Automatic calculation mode.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to load (or create) a workbook, force FormulaSettings.CalculationMode to Automatic before and after merging the range A1:B2 on the first worksheet, and save the file so that all existing formulas stay intact.
    public class PreserveFormulasDuringMerge
    {
        public static void Main(string[] args)
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
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify input file exists; if not, create a new workbook as a fallback
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
                // Optionally add a sample worksheet with data/formulas here
            }

            // Ensure calculation mode is Automatic before merging
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Perform cell merging on the first worksheet (merge A1:B2)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.Merge(0, 0, 2, 2); // firstRow, firstColumn, totalRows, totalColumns

            // Re‑assert Automatic calculation mode after merging
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
