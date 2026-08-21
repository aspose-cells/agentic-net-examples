// Title: Aspose.Cells for .NET: Replace RAND() Formulas with a Fixed Seed Value in Excel Workbooks
// Description: Load an Excel file with Aspose.Cells, scan every worksheet and used cell, detect formulas that contain the RAND() function, substitute each occurrence with a predefined static number, recalculate the workbook to lock in deterministic results, and save the updated file.
// Keywords: Aspose.Cells replace RAND | C# static seed for RAND() | deterministic Excel formulas .NET | iterate worksheets Aspose.Cells | modify cell formulas programmatically | recalculate workbook after formula change | Excel volatile function replacement | GitHub example Aspose.Cells RAND
// Common Searches: replace RAND() with constant using Aspose.Cells C# | make Excel RAND() deterministic in .NET | how to change all formulas in a workbook with Aspose.Cells | recalculate Excel after modifying formulas Aspose | static seed for RAND() in C# Excel automation
// Developer Intent: Swap every RAND() formula in an Excel workbook for a fixed numeric value to achieve repeatable calculations.
// Use Cases: Ensure financial models produce identical results across runs by fixing RAND() outputs before distribution. | Create reproducible test data sets for automated testing pipelines that generate Excel reports. | Archive legacy spreadsheets with volatile formulas converted to static values for version control compliance.
// AI Prompts: Write C# code with Aspose.Cells that finds all RAND() functions in a workbook, replaces them with a user‑defined constant, recalculates the sheet, and saves the file. | Explain step‑by‑step how to iterate through cells in Aspose.Cells, detect volatile RAND() formulas, substitute a static seed, and trigger formula recalculation. | Suggest robust error‑handling and logging enhancements for the provided RAND() replacement example.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load an Excel file with Aspose.Cells, scan every worksheet and used cell, detect formulas that contain the RAND() function, substitute each occurrence with a predefined static number, recalculate the workbook to lock in deterministic results, and save the updated file.
    public class ReplaceRandWithStaticValue
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string staticValue = "0.12345";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate over all used cells
                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                        {
                            // Replace RAND() occurrences with the static value
                            if (cell.Formula.IndexOf("RAND()", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                string newFormula = cell.Formula.Replace("RAND()", staticValue, StringComparison.OrdinalIgnoreCase);
                                cell.Formula = newFormula;
                            }
                        }
                    }
                }

                // Recalculate formulas after modifications
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ReplaceRandWithStaticValue.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
