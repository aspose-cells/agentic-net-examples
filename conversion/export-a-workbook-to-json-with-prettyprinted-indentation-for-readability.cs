// Title: How to export an Aspose.Cells workbook to a pretty‑printed JSON file with custom 4‑space indentation in C#
// AI Prompts: Generate C# code that creates a workbook, populates it with sample data, and saves it as a JSON file using Aspose.Cells with a four‑space indent. | Show how to configure JsonSaveOptions.Indent to four spaces and enable HasHeaderRow when exporting a worksheet to JSON with Aspose.Cells. | Provide a complete example that writes rows to a worksheet and writes the workbook to a formatted JSON document using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# export worksheet to JSON with indentation | JsonSaveOptions indent property usage in Aspose.Cells .NET | Save Excel data as pretty printed JSON using Aspose.Cells library
// Tags: Aspose.Cells JSON output indentation | Aspose.Cells workbook to JSON conversion | C# pretty JSON generation from Excel | JsonSaveOptions header row option | custom JSON formatting with Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, fills cells A1‑B3 with names and ages, sets JsonSaveOptions to use a four‑space indent and treat the first row as a header, and then saves the workbook as 'pretty_output.json', producing a human‑readable, indented JSON file.
class ExportWorkbookToJson
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Jane");
        worksheet.Cells["B3"].PutValue(25);

        // Set up JSON save options with pretty‑printed indentation (4 spaces)
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            Indent = "    ",          // four spaces for indentation
            HasHeaderRow = true      // treat first row as header (optional)
        };

        // Save the workbook as a JSON file using the configured options
        string outputPath = "pretty_output.json";
        workbook.Save(outputPath, saveOptions);
    }
}
