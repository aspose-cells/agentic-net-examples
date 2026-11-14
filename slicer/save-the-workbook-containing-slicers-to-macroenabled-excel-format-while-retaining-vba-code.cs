// Title: Convert an XLSX workbook with slicers to a macro‑enabled XLSM file while preserving VBA using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells save workbook with slicers as .xlsm preserving macros | How to keep VBA code when converting XLSX to XLSM with Aspose.Cells | Load an Excel file containing slicers and export to macro‑enabled format using Aspose.Cells .NET | Aspose.Cells OoxmlSaveOptions macro-enabled workbook example C# | Convert slicer‑enabled workbook to .xlsm without losing VBA project in C#
// Tags: macro-enabled XLSM export with Aspose.Cells | retain VBA project during format conversion | load slicer‑containing workbook via LoadOptions | OoxmlSaveOptions usage for macro‑enabled files | C# conversion of XLSX to XLSM preserving slicers

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the source XLSX file, loads it with LoadOptions (which also loads any embedded VBA project), and then saves it as a macro‑enabled XLSM file using OoxmlSaveOptions. All slicers and VBA code remain intact, and basic error handling reports success or failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsm";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (VBA project will be loaded automatically if present).
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(inputPath, loadOptions);

            // Save as macro‑enabled workbook. VBA project is preserved by default.
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsm);
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
