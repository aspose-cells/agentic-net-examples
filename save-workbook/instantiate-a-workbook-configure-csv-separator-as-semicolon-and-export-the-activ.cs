using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Access the active worksheet (first sheet by default)
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Ensure the first worksheet is the active one (optional)
        workbook.Worksheets.ActiveSheetIndex = 0;

        // Create CSV save options with semicolon as the delimiter
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.Separator = ';';          // Set semicolon separator
        csvOptions.ExportAllSheets = false;  // Export only the active sheet (default)

        // Save the active worksheet to a CSV file
        workbook.Save("active_sheet.csv", csvOptions);
    }
}