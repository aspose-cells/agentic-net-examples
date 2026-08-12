// Title: Skip empty rows when exporting Excel to JSON with Aspose.Cells for .NET (C#)
// Description: Learn how to export a worksheet to JSON while omitting blank rows by configuring JsonSaveOptions.IncludeEmptyRows = false (or SkipEmptyRows = true) in Aspose.Cells for .NET.
// Keywords: Aspose.Cells JSON export | IncludeEmptyRows false | SkipEmptyRows true | C# export Excel to JSON | remove empty rows JSON | .NET Aspose.Cells example | Excel to JSON without blanks | Aspose.Cells JsonSaveOptions
// Common Searches: Aspose.Cells exclude empty rows JSON | JsonSaveOptions IncludeEmptyRows false C# | How to skip blank rows when converting Excel to JSON | Export Excel worksheet to compact JSON Aspose.Cells | C# Aspose.Cells JsonSaveOptions example
// Developer Intent: Generate a JSON file from an Excel workbook that contains no empty‑row entries.
// Use Cases: Create a clean JSON payload for a REST API from a financial spreadsheet. | Produce configuration files from Excel templates without placeholder rows. | Deliver compact data sets to front‑end applications where gaps cause parsing errors.
// AI Prompts: Show a C# code sample that saves an Aspose.Cells workbook to JSON with IncludeEmptyRows set to false. | Explain how SkipEmptyRows and IncludeEmptyRows differ in Aspose.Cells JSON export. | Provide a step‑by‑step guide to verify that the resulting JSON contains no empty row objects.

using System;
using Aspose.Cells;

namespace AsposeCellsJsonExample
{
    // Learn how to export a worksheet to JSON while omitting blank rows by configuring JsonSaveOptions.IncludeEmptyRows = false (or SkipEmptyRows = true) in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data with an empty row (row index 2)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");
            // Row 3 (index 2) is left empty intentionally
            sheet.Cells["A4"].PutValue("Data3");
            sheet.Cells["B4"].PutValue("Data4");

            // Configure JSON save options to skip empty rows
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // When SkipEmptyRows is true, empty rows are excluded from the output.
                SkipEmptyRows = true
            };

            // Save the workbook as JSON using the configured options
            string outputPath = "output.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"JSON saved to '{outputPath}' with empty rows excluded.");
        }
    }
}
