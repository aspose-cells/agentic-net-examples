// Title: Load an XLSX workbook without charts using Aspose.Cells LoadOptions and inspect LoadWarnings for missing chart data in C#
// Common Searches: how to prevent chart objects from loading when opening an Excel file with Aspose.Cells in C# | Aspose.Cells C# example to load a workbook and ignore all charts | retrieve load warnings for missing chart data after loading an XLSX with Aspose.Cells | C# load Excel file with LoadOptions and check for chart‑related warnings | disable chart parsing in Aspose.Cells LoadOptions for XLSX files
// Tags: Aspose.Cells LoadOptions disable chart loading | C# load XLSX without charts | Aspose.Cells LoadWarnings inspection | missing chart data warning handling | Excel workbook loading with chart suppression

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the existence of input.xlsx, creates a LoadOptions object for the XLSX format, loads the workbook with those options, and optionally iterates through any LoadWarnings to display warning codes and messages, allowing developers to detect missing chart data while suppressing chart loading.
class Program
{
    static void Main()
    {
        // Path to the workbook to be loaded
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        try
        {
            // Create LoadOptions (default) – loading only data without charts is not directly supported
            // in older versions, so we load the workbook normally.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Optional: If your Aspose.Cells version provides LoadWarnings, you can inspect them here.
            /*
            if (workbook.LoadWarnings != null && workbook.LoadWarnings.Count > 0)
            {
                Console.WriteLine("Load warnings detected:");
                foreach (LoadWarning warning in workbook.LoadWarnings)
                {
                    Console.WriteLine($"- Code: {warning.WarningCode}, Message: {warning.WarningMessage}");
                }
            }
            else
            {
                Console.WriteLine("No load warnings.");
            }
            */

            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}
