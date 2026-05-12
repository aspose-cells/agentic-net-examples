using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add data with an intentional empty row (row 3)
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue("Data1");
        worksheet.Cells["B2"].PutValue("Data2");
        // Row 3 is left empty
        worksheet.Cells["A4"].PutValue("Data3");
        worksheet.Cells["B4"].PutValue("Data4");

        // Configure JSON save options to skip empty rows
        JsonSaveOptions saveOptions = new JsonSaveOptions();
        saveOptions.SkipEmptyRows = true; // Exclude empty rows from the JSON output

        // Save the workbook as a JSON file
        workbook.Save("output.json", saveOptions);
    }
}