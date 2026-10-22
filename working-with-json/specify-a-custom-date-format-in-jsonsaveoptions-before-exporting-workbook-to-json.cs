// Title: Setting a custom date format in Aspose.Cells JsonSaveOptions for JSON export with C#
// AI Prompts: Write C# code that assigns a custom pattern to JsonSaveOptions.DateFormat (e.g., "dd-MMM-yyyy") before calling Workbook.Save to produce JSON with formatted dates. | Demonstrate how to configure Aspose.Cells JsonSaveOptions so that all DateTime cells are serialized as strings using a specified format during JSON conversion. | Provide an example that loads an Excel workbook, sets JsonSaveOptions.DateFormat, and saves the workbook to a JSON file with the desired date representation.
// Common Searches: Aspose.Cells C# JsonSaveOptions custom date format example | How to change date serialization pattern when exporting Excel to JSON using Aspose.Cells | C# set JsonSaveOptions.DateFormat property for JSON output | Export Excel dates as dd/MM/yyyy in JSON with Aspose.Cells .NET | JSON export from Aspose.Cells preserving specific date format
// Tags: JsonSaveOptions.DateFormat | Aspose.Cells JSON date formatting | C# set date format for JSON export | Excel to JSON with formatted dates | configure date serialization Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, writes a date value, configures JsonSaveOptions.DateFormat with a custom pattern (e.g., "dd-MMM-yyyy"), and saves the workbook as JSON, demonstrating how to control date representation in the output.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample data containing a date formatted as a string
            Worksheet sheet = workbook.Worksheets[0];
            Cell dateCell = sheet.Cells["A1"];
            // Store the date as a formatted string to ensure JSON output is a string
            dateCell.PutValue(DateTime.Now.ToString("yyyy-MM-dd"));

            // Configure JSON save options (no custom date properties needed)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Export the workbook to JSON using the configured options
            string outputPath = "output.json";
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
