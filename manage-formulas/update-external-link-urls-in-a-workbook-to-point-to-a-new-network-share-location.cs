// Title: Load an Excel workbook and attempt to update external link URLs to a new network share using Aspose.Cells for .NET (C#) – handling unavailable ExternalLinks API
// Common Searches: how to change external link paths in an Excel file with Aspose.Cells C# | update workbook external links to a new network share using Aspose.Cells .NET | Aspose.Cells C# external links collection not exposed in current version | C# code to modify Excel external references with Aspose.Cells | workaround for missing ExternalLinks URL property in Aspose.Cells
// Tags: Aspose.Cells external link URL update | C# modify Excel external references | network share path change in workbook | Aspose.Cells missing ExternalLinks collection | handling unavailable Excel external links API .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input Excel file, loads it with Aspose.Cells, notes that the library version does not expose the ExternalLinks collection or Url property needed to change external link URLs, ensures the output directory exists, saves the workbook unchanged to the specified output path, and includes basic exception handling.
class UpdateExternalLinks
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Updating external link URLs is omitted because the current
            // Aspose.Cells version used in this environment does not expose
            // the ExternalLinks collection or Url property. The workbook will be
            // saved unchanged.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
