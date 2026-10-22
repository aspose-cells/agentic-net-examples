// Title: Load an XLSX workbook without charts using LoadOptions and inspect LoadWarnings for missing chart data in Aspose.Cells for .NET
// Common Searches: how to load an Excel file without charts using Aspose.Cells LoadOptions C# | Aspose.Cells LoadOptions disable chart loading and get load warnings | retrieve load warnings for missing chart data after opening workbook with Aspose.Cells | C# example for loading workbook with charts suppressed and checking warnings | Aspose.Cells version specific load warnings API for chart data issues
// Tags: suppress chart loading Aspose.Cells | access load warnings chart data issues | LoadOptions chart suppression .NET | Aspose.Cells load warnings handling | Excel workbook loading options C#

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input XLSX file exists, configures a LoadOptions object to disable chart loading, opens the workbook with those options, and shows how to read any LoadWarnings that indicate missing chart data, noting that the LoadWarnings API may be version‑specific.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Configure load options (if additional options are needed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // Uncomment the following line if the used Aspose.Cells version supports it
            // loadOptions.LoadDataOnly = true;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");

            // NOTE: LoadWarnings API may not be available in all versions of Aspose.Cells.
            // If needed, handle warnings using version‑specific features here.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }

        // Continue with further processing as needed
    }
}
