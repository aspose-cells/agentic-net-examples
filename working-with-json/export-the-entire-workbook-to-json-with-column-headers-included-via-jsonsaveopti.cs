using System;
using Aspose.Cells;

class ExportWorkbookToJson
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data with a header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);

        // Configure JSON save options to include the header row
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,
            Indent = "  " // optional indentation for readability
        };

        // Save the entire workbook as a JSON file using the options
        workbook.Save("WorkbookExport.json", jsonOptions);
    }
}