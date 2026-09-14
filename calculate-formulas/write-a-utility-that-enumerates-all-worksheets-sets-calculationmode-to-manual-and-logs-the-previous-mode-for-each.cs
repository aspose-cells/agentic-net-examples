// Title: C# console utility to list worksheets, log the workbook’s current calculation mode, and switch to Manual calculation using Aspose.Cells
// AI Prompts: Write a C# program that loads an Excel file with Aspose.Cells, iterates through all worksheets, prints the workbook’s existing FormulaSettings.CalculationMode, then sets the calculation mode to Manual and saves the file. | Show how to capture the previous CalcModeType value from a workbook’s FormulaSettings, output it per worksheet, and update the setting to Manual in Aspose.Cells for .NET. | Create a command‑line tool that accepts input and output paths, reads the workbook, logs the current calculation mode for each sheet, changes the mode to Manual, and writes the updated workbook.
// Common Searches: Aspose.Cells C# enumerate worksheets and display current calculation mode before setting manual | How to log previous CalcModeType for each sheet in an Excel workbook using Aspose.Cells .NET | Command line program to change Aspose.Cells workbook calculation mode to Manual while preserving original mode
// Tags: iterate worksheets read calculation mode Aspose.Cells | set workbook formula calculation mode manual Aspose.Cells | log previous CalcModeType per sheet C# | console application load save workbook Aspose.Cells | command line utility change calculation settings .NET

using System;
using Aspose.Cells;

namespace AsposeCellsUtility
{
    // Loads a workbook, prints the current FormulaSettings.CalculationMode for each worksheet, switches the workbook’s calculation mode to Manual, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input workbook path and output workbook path
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: AsposeCellsUtility <inputPath> <outputPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the current calculation mode (applies to the whole workbook)
            CalcModeType previousMode = workbook.Settings.FormulaSettings.CalculationMode;

            // Enumerate all worksheets and log the previous calculation mode for each
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\": Previous CalculationMode = {previousMode}");
            }

            // Set the calculation mode to Manual for the workbook
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Save the modified workbook to the specified output path
            workbook.Save(outputPath);
        }
    }
}
