// Title: Export an Excel table to indented JSON with header keys using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a header row (Name, Age, City) and sample data, then uses Aspose.Cells JsonSaveOptions (HasHeaderRow = true, Indent = " ") to save the worksheet as a pretty‑printed JSON file where column headers become object properties.
// Keywords: Aspose.Cells | C# | Export to JSON | JsonSaveOptions | header row as keys | JSON indentation | pretty printed JSON | Excel to JSON | Workbook.Save | ExportTableToJson | GitHub example | code sample
// Common Searches: Aspose.Cells export worksheet to JSON with headers | JsonSaveOptions indent spaces C# | How to save Excel as pretty printed JSON using Aspose.Cells | C# export table to JSON file with column names | Aspose.Cells example for JSON output formatting
// Developer Intent: Generate a readable JSON file from an Excel worksheet where the first row supplies property names and the output is indented for easy version‑control diff.
// Use Cases: Provide API clients with a clean JSON payload generated directly from spreadsheet data. | Create configuration or seed data files from Excel that remain human‑readable in source control. | Export test datasets from in‑memory workbooks for automated testing or CI pipelines.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet range to a JSON file using 4‑space indentation and include the first row as property names. | Show how to configure JsonSaveOptions for pretty‑printed JSON output when saving a workbook in .NET. | Explain error handling best practices when calling Workbook.Save with JsonSaveOptions in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsJsonExportDemo
{
    // This example creates a workbook, adds a header row (Name, Age, City) and sample data, then uses Aspose.Cells JsonSaveOptions (HasHeaderRow = true, Indent = " ") to save the worksheet as a pretty‑printed JSON file where column headers become object properties.
    public class ExportTableToJson
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate the worksheet with a header row and some data rows
                worksheet.Cells["A1"].PutValue("Name");
                worksheet.Cells["B1"].PutValue("Age");
                worksheet.Cells["C1"].PutValue("City");

                worksheet.Cells["A2"].PutValue("John");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["C2"].PutValue("New York");

                worksheet.Cells["A3"].PutValue("Alice");
                worksheet.Cells["B3"].PutValue(25);
                worksheet.Cells["C3"].PutValue("London");

                worksheet.Cells["A4"].PutValue("Bob");
                worksheet.Cells["B4"].PutValue(28);
                worksheet.Cells["C4"].PutValue("Paris");

                // Configure JSON save options:
                // - Include the header row as keys (HasHeaderRow = true)
                // - Use 4 spaces for indentation (Indent = "    ")
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,
                    Indent = "    " // 4 spaces
                };

                // Save the workbook as a JSON file using the configured options
                string outputPath = "ExportedTable.json";
                workbook.Save(outputPath, jsonOptions);

                Console.WriteLine($"JSON file saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportTableToJson.Run();
        }
    }
}
