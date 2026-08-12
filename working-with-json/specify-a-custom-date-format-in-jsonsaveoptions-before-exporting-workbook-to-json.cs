// Title: C# – Export Aspose.Cells Workbook to JSON with a Custom Date Format Using JsonSaveOptions
// Description: Demonstrates how to export a workbook to JSON while forcing dates to appear in a chosen pattern (e.g., dd‑MM‑yyyy). The example creates a workbook, writes a date as a formatted string, uses JsonSaveOptions (which lacks a direct date‑format setting), and saves the file.
// Keywords: Aspose.Cells | C# | JsonSaveOptions | custom date format | JSON export | date string formatting | Workbook to JSON | Aspose.Cells JSON | date representation
// Common Searches: Aspose.Cells export to JSON with specific date format C# | How to change date pattern when saving workbook as JSON | JsonSaveOptions custom date format example | C# Aspose.Cells JSON date formatting dd-MM-yyyy | Set date format for JSON output in Aspose.Cells
// Developer Intent: Define the appearance of date values in the JSON file produced from an Aspose.Cells workbook.
// Use Cases: Integrate with a downstream API that expects dates in dd‑MM‑yyyy format. | Maintain consistent date strings across multiple exported JSON files without altering cell types. | Show that JsonSaveOptions does not expose a date‑format property, so cell‑level formatting is required.
// AI Prompts: Generate C# code that exports an Aspose.Cells workbook to JSON with dates formatted as "yyyy/MM/dd" while keeping original cell types unchanged. | Explain why JsonSaveOptions lacks a direct date‑format option and how to work around it for custom JSON date strings. | Provide a method to apply a uniform date format to all date cells in a workbook before JSON export using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExample
{
    // Demonstrates how to export a workbook to JSON while forcing dates to appear in a chosen pattern (e.g., dd‑MM‑yyyy). The example creates a workbook, writes a date as a formatted string, uses JsonSaveOptions (which lacks a direct date‑format setting), and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add header
                sheet.Cells["A1"].PutValue("Date");

                // Add date value formatted as a string to achieve custom JSON date format
                DateTime date = new DateTime(2023, 5, 15);
                sheet.Cells["B1"].PutValue(date.ToString("dd-MM-yyyy"));

                // Configure JSON save options (no custom date format property needed)
                JsonSaveOptions saveOptions = new JsonSaveOptions();

                // Export the workbook to a JSON file
                string outputPath = "ExportedWithCustomDateFormat.json";
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
