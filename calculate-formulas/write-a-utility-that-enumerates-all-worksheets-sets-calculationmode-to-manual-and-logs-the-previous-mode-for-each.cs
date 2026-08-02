// Title: List All Worksheets, Log Existing Calculation Mode, and Set Workbook to Manual with Aspose.Cells for .NET
// Description: Loads an Excel workbook, reads the current CalcModeType, writes the mode for each worksheet to the console, switches the calculation mode to Manual, and saves the modified file.
// Keywords: Aspose.Cells | C# | CalcModeType | manual calculation mode | enumerate worksheets | log calculation mode | workbook formula settings | performance optimization | Excel automation | save workbook
// Common Searches: Aspose.Cells set calculation mode to manual | how to get current calculation mode in Aspose.Cells | list all worksheets and their formula settings Aspose.Cells | log workbook calculation mode before changing it | C# change Excel calculation mode with Aspose.Cells
// Developer Intent: Read the workbook's current calculation mode, output it for each worksheet, change the mode to Manual, and persist the change.
// Use Cases: Create a diagnostic log of the original calculation mode before disabling automatic recalculation for large data imports. | Improve batch‑processing speed by recording the existing mode, switching to Manual, and later restoring the original setting. | Audit multiple workbooks to ensure they all use Manual calculation during server‑side processing.
// AI Prompts: Generate C# code using Aspose.Cells that iterates through every worksheet, prints the workbook's current CalcModeType, sets the calculation mode to Manual, and saves the file. | Explain how to capture the original CalcModeType, apply Manual calculation for intensive updates, and revert to the saved mode after processing with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsUtility
{
    // Loads an Excel workbook, reads the current CalcModeType, writes the mode for each worksheet to the console, switches the calculation mode to Manual, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with actual file path)
            string inputPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the current calculation mode (applies to the whole workbook)
            CalcModeType previousMode = workbook.Settings.FormulaSettings.CalculationMode;

            // Enumerate all worksheets and log the previous calculation mode for each
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\": previous calculation mode = {previousMode}");
            }

            // Set the calculation mode to Manual for the workbook
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Save the modified workbook (replace with desired output path)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}
