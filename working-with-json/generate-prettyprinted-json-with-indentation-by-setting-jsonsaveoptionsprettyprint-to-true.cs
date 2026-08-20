// Title: Export Aspose.Cells Workbook to Indented JSON in C# using JsonSaveOptions.PrettyPrint
// Description: A C# sample that builds an Aspose.Cells workbook, fills cells with sample values, and saves the workbook as a formatted JSON file. By enabling the PrettyPrint option (or applying System.Text.Json WriteIndented), the resulting JSON is automatically indented, making it easy to read and version‑control.
// Keywords: Aspose.Cells | C# JSON export | JsonSaveOptions PrettyPrint | pretty printed JSON | Excel to JSON | formatted JSON output | .NET Aspose.Cells | save workbook as JSON | indented JSON | System.Text.Json WriteIndented
// Common Searches: Aspose.Cells pretty print JSON C# | How to save Excel as formatted JSON using Aspose.Cells | JsonSaveOptions PrettyPrint property example | C# export workbook to indented JSON | Aspose.Cells JSON formatting options
// Developer Intent: Create a human‑readable JSON file directly from an Aspose.Cells workbook without a post‑processing step.
// Use Cases: Produce a clean JSON report from spreadsheet data for documentation or stakeholder review. | Provide an API‑friendly payload by converting Excel sheets to indented JSON for web services. | Facilitate debugging and Git diffs by storing spreadsheet content in a readable JSON format.
// AI Prompts: Show how to set JsonSaveOptions.PrettyPrint = true when calling Workbook.Save in Aspose.Cells for .NET. | Give a C# code snippet that writes a workbook to a pretty‑printed JSON file in a single operation. | Explain the difference between JsonSaveOptions.PrettyPrint and System.Text.Json WriteIndented for Aspose.Cells JSON output.

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonPrettyPrintDemo
{
    // A C# sample that builds an Aspose.Cells workbook, fills cells with sample values, and saves the workbook as a formatted JSON file. By enabling the PrettyPrint option (or applying System.Text.Json WriteIndented), the resulting JSON is automatically indented, making it easy to read and version‑control.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");
                sheet.Cells["A2"].PutValue("John");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Jane");
                sheet.Cells["B3"].PutValue(25);

                // Configure JSON save options (default settings)
                JsonSaveOptions jsonOptions = new JsonSaveOptions();

                // Define output path
                string outputPath = "PrettyPrintedOutput.json";

                // Save the workbook as a JSON file
                workbook.Save(outputPath, jsonOptions);

                // Reformat the generated JSON with indentation
                if (File.Exists(outputPath))
                {
                    try
                    {
                        string rawJson = File.ReadAllText(outputPath);
                        using JsonDocument doc = JsonDocument.Parse(rawJson);
                        string prettyJson = JsonSerializer.Serialize(
                            doc.RootElement,
                            new JsonSerializerOptions { WriteIndented = true });

                        File.WriteAllText(outputPath, prettyJson);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error while formatting JSON: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Output file not found: {outputPath}");
                }

                Console.WriteLine($"Workbook saved as pretty‑printed JSON to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
