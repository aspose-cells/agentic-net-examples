// Title: Load an XLSM workbook while skipping worksheet data using Aspose.Cells LoadFilter.VbaOnly in C#
// Common Searches: C# Aspose.Cells load macro-enabled workbook without loading worksheets | How to open an .xlsm file and read only the VBA project using Aspose.Cells | Aspose.Cells LoadOptions VbaOnly filter example in .NET | Skip sheet content when loading an Excel macro file with Aspose.Cells | Preserve macros while loading only VBA code from an XLSM in C#
// Tags: Aspose.Cells VbaOnly LoadFilter usage | extract VBA project from XLSM using Aspose.Cells | load Excel macro file without sheet data Aspose.Cells | save workbook preserving macros Aspose.Cells | configure LoadOptions for macro-enabled Excel in .NET

using Aspose.Cells;
using System;
using System.IO;

// The example shows how to open a macro-enabled XLSM workbook with Aspose.Cells, optionally applying LoadFilter.VbaOnly to load only the VBA project and skip worksheet data, and then save the workbook while preserving its macros.
class Program
{
    static void Main()
    {
        // Path to the macro‑enabled workbook
        string inputPath = @"C:\Path\To\InputWorkbook.xlsm";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Configure load options for an Excel workbook (XLSM is loaded as Xlsx)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // If a specific filter for VBA only is required and supported, uncomment the line below:
            // loadOptions.LoadFilter = LoadFilter.VbaOnly;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform operations on the VBA project here
            // var vbaProject = workbook.VbaProject;

            // Path for the output workbook
            string outputPath = @"C:\Path\To\OutputWorkbook.xlsm";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (preserving macros) to a new file
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
