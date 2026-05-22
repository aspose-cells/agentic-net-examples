using System;
using Aspose.Cells;

class ExportTableToJson
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate header row
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["C1"].PutValue("City");

        // Populate data rows
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["C2"].PutValue("New York");

        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);
        worksheet.Cells["C3"].PutValue("London");

        // Set JSON export options: include headers and format with indentation
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,
            Indent = "    " // 4 spaces for readability
        };

        // Export the workbook (entire sheet) to a formatted JSON file
        workbook.Save("TableExport.json", jsonOptions);
    }
}