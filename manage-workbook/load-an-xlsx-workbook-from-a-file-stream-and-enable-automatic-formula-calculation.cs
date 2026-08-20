// Title: Load XLSX from FileStream with Automatic Formula Calculation using Aspose.Cells for .NET (C#)
// Description: Opens an XLSX file via a read‑only FileStream, uses LoadOptions with ParsingFormulaOnOpen, sets the workbook to Automatic calculation mode, optionally forces a recalculation, and saves the result.
// Keywords: Aspose.Cells | C# | LoadOptions | ParsingFormulaOnOpen | automatic calculation | Workbook | FileStream | calculate formulas | save workbook | XLSX
// Common Searches: Aspose.Cells load workbook from filestream | enable automatic formula calculation Aspose.Cells .NET | ParsingFormulaOnOpen example C# | recalculate formulas after loading workbook Aspose.Cells | how to set calculation mode to Automatic Aspose.Cells
// Developer Intent: Load an XLSX workbook from a stream, ensure formulas are parsed, enable automatic calculation, optionally recalc immediately, and save the updated file.
// Use Cases: Web API that receives uploaded Excel files, recalculates all formulas automatically, and returns the processed workbook. | Batch processing of multiple templates where each workbook must be loaded, formulas evaluated, and saved as a final report. | Generating dynamic reports by loading a template, triggering formula evaluation, and exporting the completed file.
// AI Prompts: Generate C# code that loads an Excel workbook from a MemoryStream with Aspose.Cells, enables automatic calculation, and returns the workbook as a byte array. | Show how to configure LoadOptions to ignore formula parsing errors while still allowing automatic formula calculation in Aspose.Cells for .NET. | Explain how to switch a workbook's calculation mode to Manual, recalculate specific cells, and then revert to Automatic using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Opens an XLSX file via a read‑only FileStream, uses LoadOptions with ParsingFormulaOnOpen, sets the workbook to Automatic calculation mode, optionally forces a recalculation, and saves the result.
class LoadWorkbookWithAutoCalc
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Open the file as a read‑only stream
            using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                // Create load options and ensure formulas are parsed when the workbook is opened
                LoadOptions loadOptions = new LoadOptions
                {
                    ParsingFormulaOnOpen = true // parse formulas during load
                };

                // Load the workbook from the stream with the specified options
                Workbook workbook = new Workbook(stream, loadOptions);

                // Enable automatic formula calculation (default mode)
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                // Optionally calculate formulas immediately
                workbook.CalculateFormula();

                // Save the workbook to a new file
                workbook.Save("output.xlsx", SaveFormat.Xlsx);
            }

            Console.WriteLine("Workbook processed and saved as \"output.xlsx\" successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
