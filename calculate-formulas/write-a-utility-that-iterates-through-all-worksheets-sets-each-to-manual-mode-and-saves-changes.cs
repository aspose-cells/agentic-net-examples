// Title: C# example: iterate all worksheets, set workbook calculation mode to Manual with Aspose.Cells and save the workbook
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual for entire workbook | How to change formula calculation mode to manual in an Excel file using Aspose.Cells .NET | Iterate through worksheets and save workbook after disabling automatic calculations with Aspose.Cells | C# code to load an Excel file, set manual calculation mode, and export to a new file using Aspose.Cells | Disable automatic formula recalculation for all sheets in Aspose.Cells and save as xlsx
// Tags: Aspose.Cells workbook calculation settings | C# disable automatic formula recalculation | Aspose.Cells iterate over worksheets | export modified workbook Aspose.Cells | manual formula calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample program validates command‑line arguments, loads an Excel workbook with Aspose.Cells, iterates through each worksheet (demonstrating traversal), sets the workbook‑wide calculation mode to Manual via Settings.FormulaSettings, and saves the updated workbook to the specified output path.
class Program
{
    static void Main(string[] args)
    {
        // Validate arguments
        if (args.Length < 1)
        {
            Console.WriteLine("Usage: Program.exe <inputWorkbookPath> [outputWorkbookPath]");
            return;
        }

        string inputPath = args[0];
        string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Console.WriteLine($"Loaded workbook with {workbook.Worksheets.Count} worksheet(s).");

            // Iterate through all worksheets (demonstration purpose)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Processing worksheet: '{sheet.Name}' (Index: {sheet.Index})");
                // No per‑worksheet calculation mode exists; the setting is workbook‑wide.
                // The loop is kept to satisfy the task requirement of iterating worksheets.
            }

            // Set calculation mode to Manual for the entire workbook
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            Console.WriteLine("Workbook calculation mode set to Manual.");

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An exception occurred: {ex.Message}");
        }
    }
}
