// Title: Why JsonSaveOptions cannot apply a custom date format when exporting an Aspose.Cells workbook to JSON (C#)
// AI Prompts: Generate C# code that attempts to set a date pattern in JsonSaveOptions, captures the resulting limitation, and proposes a workaround for formatting dates in the exported JSON. | Explain the limitation of JsonSaveOptions regarding date formatting and describe alternative approaches to control date representation in JSON output using Aspose.Cells. | Provide a step‑by‑step guide for post‑processing the JSON file produced by Aspose.Cells to replace default date strings with a custom format.
// Common Searches: Aspose.Cells JsonSaveOptions does not support custom date format | How to change date format in JSON output when saving workbook with Aspose.Cells .NET | Workaround for formatting dates in Aspose.Cells JSON export C# | Export workbook to JSON with specific date pattern using Aspose.Cells
// Tags: json export custom date format Aspose.Cells | JsonSaveOptions date formatting limitation | Aspose.Cells workbook to JSON date serialization | C# Aspose.Cells JSON export settings | post‑process JSON dates Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Json;

// The sample creates a workbook, inserts a DateTime value, and attempts to configure JsonSaveOptions for a custom date format. Because JsonSaveOptions lacks a DateFormat property, the export uses the default date representation. The example highlights this limitation and suggests alternative strategies such as cell formatting before export or post‑processing the generated JSON to achieve the desired date pattern.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add a worksheet and some data with a date value
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 5, 17));

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions();
            // JsonSaveOptions in this version does not expose a DateFormat property.
            // The default date format will be used for JSON export.

            string outputFile = "ExportedData.json";

            // Export the workbook to JSON using the configured options
            workbook.Save(outputFile, jsonOptions);
            Console.WriteLine($"Workbook successfully exported to JSON: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
