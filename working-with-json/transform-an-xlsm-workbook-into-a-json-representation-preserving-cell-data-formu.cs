using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSM workbook
        string inputPath = "input.xlsm";

        // Path where the JSON representation will be saved
        string outputPath = "output.json";

        // Load the workbook (preserves macros, formulas, etc.)
        Workbook workbook = new Workbook(inputPath);

        // Configure JSON save options to export the full Excel structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Preserve cell values, formulas, and worksheet hierarchy
            ToExcelStruct = true
        };

        // Save the workbook as JSON using the configured options
        workbook.Save(outputPath, jsonOptions);

        Console.WriteLine("Workbook successfully converted to JSON.");
    }
}