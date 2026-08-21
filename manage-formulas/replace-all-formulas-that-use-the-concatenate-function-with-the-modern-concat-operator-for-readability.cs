// Title: Convert CONCATENATE to CONCAT in Excel using Aspose.Cells for .NET
// Description: Loads an Excel workbook, scans every used cell, detects formulas that contain the legacy CONCATENATE function, replaces it with the modern CONCAT operator, recalculates the workbook, and saves the updated file.
// Keywords: Aspose.Cells | C# Excel formula update | CONCATENATE to CONCAT conversion | bulk formula replacement | .NET Excel modernization | replace legacy Excel functions | Excel 365 compatibility
// Common Searches: Aspose.Cells replace CONCATENATE with CONCAT | C# bulk update Excel formulas | Convert old CONCATENATE formulas to CONCAT .NET | Recalculate workbook after formula changes Aspose | Iterate cells and modify formulas Aspose.Cells
// Developer Intent: Automatically change all CONCATENATE functions in a workbook to the CONCAT operator and save the revised file.
// Use Cases: Upgrade legacy spreadsheets to the newer CONCAT syntax required by recent Excel versions. | Run a batch job that cleans up formulas across multiple workbooks before distribution. | Ensure formula compatibility when migrating older Excel files to cloud‑based reporting platforms.
// AI Prompts: Write C# code that scans every worksheet in an Aspose.Cells workbook, replaces CONCATENATE with CONCAT in formulas, recalculates, and saves the result. | Create a method that logs the address of each cell where a CONCATENATE formula was changed to CONCAT using Aspose.Cells. | Develop a reusable Aspose.Cells utility class for bulk formula transformations, including CONCATENATE‑to‑CONCAT replacement with optional error handling.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaUpdate
{
    // Loads an Excel workbook, scans every used cell, detects formulas that contain the legacy CONCATENATE function, replaces it with the modern CONCAT operator, recalculates the workbook, and saves the updated file.
    public class ReplaceConcatenateWithConcat
    {
        public static void Run()
        {
            // Input and output file paths (replace with actual paths)
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"Input file not found: {inputPath}");
                }

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range to limit iteration
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Loop through each cell in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Process only cells that contain a formula
                            if (cell.IsFormula)
                            {
                                string formula = cell.Formula; // Formula string includes leading '='

                                // Check for the legacy CONCATENATE function (case‑insensitive)
                                if (formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // Replace the function name with the modern CONCAT operator
                                    string updatedFormula = formula.Replace("CONCATENATE", "CONCAT", StringComparison.OrdinalIgnoreCase);

                                    // Assign the new formula back to the cell
                                    cell.Formula = updatedFormula;
                                }
                            }
                        }
                    }
                }

                // Recalculate all formulas after modifications
                workbook.CalculateFormula();

                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                // Optionally rethrow or handle specific exceptions as needed
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceConcatenateWithConcat.Run();
        }
    }
}
