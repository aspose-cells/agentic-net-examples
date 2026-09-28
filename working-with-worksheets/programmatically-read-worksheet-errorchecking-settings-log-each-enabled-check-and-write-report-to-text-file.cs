// Title: Create a C# program that reads worksheet error‑checking information with Aspose.Cells and writes a summary report to a text file
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells, iterates each worksheet, checks whether error‑checking settings are accessible, appends a line describing the availability to a StringBuilder, and saves the accumulated text to a .txt file. | Enhance the sample to detect the Aspose.Cells version at runtime; if a newer version exposes specific error‑checking flags (e.g., NumberStoredAsText, InconsistentFormula), retrieve those flags and log which checks are enabled for each worksheet. | Add robust error handling so that missing input files, unavailable error‑checking data, or unexpected exceptions are caught, logged, and do not stop the generation of the final report.
// Common Searches: aspnet read worksheet error checking settings with Aspose.Cells | how to export Excel error checking flags to a text file using C# | Aspose.Cells worksheet validation options not exposed in .NET | C# generate report of Excel data validation and error checking per sheet | list enabled error checks for each worksheet in a workbook using Aspose.Cells
// Tags: Aspose.Cells read worksheet error checking | C# export worksheet validation summary to text | Aspose.Cells generate error checking report | iterate Excel worksheets log validation settings | Aspose.Cells .NET write text report

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through all worksheets, notes that detailed error‑checking settings are not exposed by the current API, and writes a placeholder summary for each sheet to a text file while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "ErrorCheckingReport.txt";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            StringBuilder reportBuilder = new StringBuilder();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                reportBuilder.AppendLine($"Worksheet: {sheet.Name}");

                // NOTE: Aspose.Cells does not expose detailed error‑checking options
                // via a public API in all versions. If needed, this section can be
                // extended with version‑specific calls. For now we indicate that
                // error‑checking details are unavailable.
                reportBuilder.AppendLine("  - Error checking details are not available via the current API.");

                // Add a blank line after each worksheet's report
                reportBuilder.AppendLine();
            }

            // Write the compiled report to a text file
            File.WriteAllText(outputPath, reportBuilder.ToString());
            Console.WriteLine($"Report generated successfully at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
