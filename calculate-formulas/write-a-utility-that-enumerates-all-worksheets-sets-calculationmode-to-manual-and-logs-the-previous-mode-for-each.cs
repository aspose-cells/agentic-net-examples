// Title: Aspose.Cells .NET – List All Worksheets, Log Their Current Calculation Mode, and Switch to Manual
// Description: Loads an Excel workbook, iterates through every worksheet, records the workbook's existing FormulaSettings.CalculationMode for each sheet, prints the information to the console, changes the mode to Manual, and saves the file.
// Keywords: Aspose.Cells calculation mode manual | log previous calculation mode C# | enumerate worksheets Aspose.Cells | FormulaSettings.CalculationMode | Aspose.Cells performance optimization | .NET Excel workbook settings
// Common Searches: how to set calculation mode to manual with Aspose.Cells | retrieve and log current calculation mode for each sheet | Aspose.Cells enumerate all worksheets example | C# change Excel formula calculation mode using Aspose | log workbook settings before modifying Aspose.Cells
// Developer Intent: Capture the existing calculation mode for each worksheet, output it for auditing, then enforce Manual calculation across the entire workbook.
// Use Cases: Speed up bulk data imports by disabling automatic recalculation while preserving the original setting for later restoration. | Create a reproducible audit trail of formula calculation settings before running batch transformations. | Standardize workbook behavior in automated reporting pipelines by forcing Manual mode after documenting the prior configuration.
// AI Prompts: Generate C# code that loops through all worksheets in an Aspose.Cells workbook, prints each sheet's current CalculationMode, and sets the workbook to Manual calculation. | Provide an Aspose.Cells .NET example that logs the previous FormulaSettings.CalculationMode for every worksheet before switching to manual mode and saving the file. | Write a C# utility using Aspose.Cells that records each worksheet's calculation mode, changes the mode to Manual, and outputs the original values to the console.

using System;
using Aspose.Cells;

namespace AsposeCellsUtility
{
    // Loads an Excel workbook, iterates through every worksheet, records the workbook's existing FormulaSettings.CalculationMode for each sheet, prints the information to the console, changes the mode to Manual, and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Retrieve the current calculation mode before changing it
                CalcModeType previousMode = workbook.Settings.FormulaSettings.CalculationMode;

                // Log the worksheet name and its previous calculation mode
                Console.WriteLine($"Worksheet '{sheet.Name}': previous calculation mode = {previousMode}");

                // Set the calculation mode to Manual
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
        }
    }
}
