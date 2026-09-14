// Title: C# console utility that sets all worksheets to Manual calculation mode with Aspose.Cells and saves the workbook
// AI Prompts: Generate a C# method that loads an Excel file using Aspose.Cells, changes the workbook's FormulaSettings.CalculationMode to Manual, and writes the updated file to a specified output path. | Create a .NET command‑line program that accepts input and output file arguments, applies Manual calculation mode to every worksheet in the workbook via Aspose.Cells, and saves the result.
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual for all sheets | How to change formula calculation mode to Manual using Aspose.Cells .NET | C# program to load Excel, set manual calculation, and save with Aspose.Cells | Iterate through worksheets and apply manual calculation mode Aspose.Cells | Command line tool for changing Excel calculation mode with Aspose.Cells
// Tags: Aspose.Cells manual calculation mode | C# set workbook calculation mode manual | Aspose.Cells iterate worksheets | save workbook after changing formula settings | command line Aspose.Cells utility

using System;
using Aspose.Cells;

// Utility that loads an Excel workbook with Aspose.Cells, switches the workbook's formula calculation mode to Manual (affecting all worksheets), and saves the modified file to a new location.
public class SetManualCalculationUtility
{
    // Sets the calculation mode of the workbook to Manual for all worksheets and saves the file.
    public static void SetManualMode(string inputFilePath, string outputFilePath)
    {
        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(inputFilePath);

        // Set the workbook's calculation mode to Manual.
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Iterate through all worksheets (optional processing per sheet).
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Example placeholder: you can add per‑sheet logic here if needed.
            // Currently no per‑sheet setting is required for calculation mode.
        }

        // Save the modified workbook to the output path.
        workbook.Save(outputFilePath);
    }

    // Example entry point.
    public static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: SetManualCalculationUtility <inputFilePath> <outputFilePath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        SetManualMode(inputPath, outputPath);
        Console.WriteLine($"Workbook saved with Manual calculation mode to '{outputPath}'.");
    }
}
