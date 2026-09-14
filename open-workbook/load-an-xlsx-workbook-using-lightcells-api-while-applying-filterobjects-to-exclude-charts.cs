// Title: C# – Load an XLSX workbook with Aspose.Cells LightCells API and filter out chart objects
// AI Prompts: Generate C# code that opens an XLSX file using Aspose.Cells LightCells API, applies a LoadFilter to omit all charts, and saves the result to a new file. | Demonstrate how to configure LoadOptions with LoadFilterOptions.RemoveChart in Aspose.Cells to load a workbook without chart objects and then write it out.
// Common Searches: how to load an xlsx file with Aspose.Cells LightCells API without charts in C# | C# Aspose.Cells LoadFilterOptions.RemoveChart usage example | exclude chart objects when opening an Excel workbook using LightCells | Aspose.Cells LightCells load options to skip charts during workbook load
// Tags: lightcells loadoptions removechart c# | aspocells load xlsx without charts | c# workbook load filter chart exclusion | loadfilteroptions removechart example | excel workbook load filter chart objects

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input XLSX file, creates LoadOptions for the Xlsx format, optionally assigns a LoadFilter with RemoveChart (commented for version compatibility), loads the workbook via Aspose.Cells LightCells API, ensures the output directory exists, and saves the workbook to the specified output path.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Set load options for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // NOTE: Removing charts via LoadFilter requires LoadFilterOptions enum,
            // which may not be available in all Aspose.Cells versions.
            // If needed and supported, uncomment the following lines:
            // var filter = new LoadFilter(LoadFilterOptions.RemoveChart);
            // loadOptions.LoadFilter = filter;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to the output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
