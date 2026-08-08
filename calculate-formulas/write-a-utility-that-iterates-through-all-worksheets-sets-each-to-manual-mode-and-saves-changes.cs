// Title: C# utility to set workbook calculation mode to Manual for all worksheets with Aspose.Cells
// Description: A command‑line program that loads an Excel file using Aspose.Cells for .NET, switches the workbook's FormulaSettings.CalculationMode to Manual, optionally iterates each worksheet to adjust properties, and saves the updated workbook to a new location.
// Keywords: Aspose.Cells manual calculation mode | C# set CalcModeType.Manual | iterate worksheets Aspose.Cells | save workbook after changing settings | disable automatic formula recalculation .NET | Aspose.Cells workbook settings
// Common Searches: Aspose.Cells set calculation mode manual C# | how to disable automatic formula calculation in Aspose.Cells | iterate all worksheets and change properties Aspose.Cells | save workbook after changing formula settings Aspose.Cells | command line tool to set manual calculation mode Aspose.Cells
// Developer Intent: Change the workbook’s formula calculation mode to Manual for the entire file and persist the change by saving to a new workbook.
// Use Cases: Improve performance when performing bulk data updates by turning off automatic recalculation. | Prepare a workbook for downstream systems that require formulas to stay unevaluated until a later step. | Ensure consistent manual calculation mode across all sheets while customizing per‑sheet properties such as selection state.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, sets the workbook calculation mode to Manual, loops through each worksheet to modify a property, and saves the result to a specified path. | Explain how CalcModeType.Manual affects formula evaluation in Aspose.Cells and whether any worksheet‑level settings can override it. | Provide a step‑by‑step command‑line example for disabling automatic calculation in a workbook using Aspose.Cells and saving it as a new file.

using System;
using Aspose.Cells;

namespace AsposeCellsUtility
{
    // A command‑line program that loads an Excel file using Aspose.Cells for .NET, switches the workbook's FormulaSettings.CalculationMode to Manual, optionally iterates each worksheet to adjust properties, and saves the updated workbook to a new location.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input file path and output file path.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: AsposeCellsUtility <inputFilePath> <outputFilePath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // Set the calculation mode of the workbook to Manual.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Iterate through all worksheets (optional, shows access to each sheet).
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Example operation: ensure the sheet is selected when opened.
                // This does not affect calculation mode but demonstrates per‑sheet handling.
                sheet.IsSelected = false;
            }

            // Save the modified workbook to the output path.
            workbook.Save(outputPath);
        }
    }
}
