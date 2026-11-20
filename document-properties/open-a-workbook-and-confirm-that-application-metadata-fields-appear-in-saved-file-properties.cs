// Title: How to set and confirm the Application built‑in document property in an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set Application built‑in document property in existing .xlsx | Read built‑in Application property from a workbook using Aspose.Cells | Check Excel file metadata after saving with Aspose.Cells .NET | Update and verify Excel document properties programmatically in C# | Persist workbook metadata when saving with Aspose.Cells
// Tags: modify Application built‑in property Aspose.Cells | retrieve built‑in document properties Aspose.Cells | Excel file metadata persistence using Aspose.Cells | update workbook built‑in properties C# | verify document property after workbook save .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx file, changes the built‑in Application document property to "Aspose.Cells", saves the workbook, reloads it, and confirms that the Application property persists, handling missing files and exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the built‑in Application property via its Value member
            var appProp = workbook.BuiltInDocumentProperties["Application"];
            if (appProp != null)
            {
                appProp.Value = "Aspose.Cells";
            }

            // Save the workbook with the updated metadata
            workbook.Save(outputPath);

            // Reload the saved workbook to verify the metadata persisted
            Workbook reloaded = new Workbook(outputPath);

            // Retrieve the Application property value
            string? appProperty = reloaded.BuiltInDocumentProperties["Application"]?.Value?.ToString();

            // Output verification result
            if (!string.IsNullOrEmpty(appProperty))
            {
                Console.WriteLine("Application metadata field found: " + appProperty);
            }
            else
            {
                Console.WriteLine("Application metadata field not found.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
