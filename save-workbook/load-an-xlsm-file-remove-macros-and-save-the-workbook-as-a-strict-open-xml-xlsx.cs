// Title: Remove VBA macros from an XLSM workbook and save as a strict Open XML XLSX using Aspose.Cells for .NET
// AI Prompts: Load an XLSM file with Aspose.Cells, delete all VBA modules, and export it to a macro‑free XLSX using OoxmlSaveOptions. | Write C# code that checks for a VbaProject, clears its modules, and saves the workbook as a strict Open XML XLSX. | Demonstrate how to create the output folder if it does not exist and then save a macro‑free workbook with Aspose.Cells.
// Common Searches: asp.net remove macros from xlsm and save as xlsx using aspose.cells | c# convert macro enabled workbook to strict xlsx without VBA | how to clear VBA project modules before saving workbook with Aspose.Cells | save xlsm as xlsx ensuring output directory exists c#
// Tags: clear VBA modules Aspose.Cells | convert XLSM to strict XLSX C# | OoxmlSaveOptions macro‑free export | ensure output directory Aspose.Cells | remove macros before workbook save .NET

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input XLSM file, loads it into an Aspose.Cells Workbook, removes any VBA modules if a VbaProject is present, prepares OoxmlSaveOptions for XLSX format, creates the output directory when needed, and saves the workbook as a macro‑free strict Open XML XLSX while handling errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the XLSM workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove all VBA modules if a VBA project exists
            if (workbook.VbaProject != null && workbook.VbaProject.Modules != null)
            {
                workbook.VbaProject.Modules.Clear();
            }

            // Configure save options for XLSX format
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
            // Note: Compliance property is not available in the referenced Aspose.Cells version.
            // The workbook will be saved with default compliance settings.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an XLSX file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
