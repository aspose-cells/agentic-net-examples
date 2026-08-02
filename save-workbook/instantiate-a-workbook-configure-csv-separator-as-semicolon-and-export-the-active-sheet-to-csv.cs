using System;
using Aspose.Cells;

class ExportActiveSheetToCsv
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the active worksheet (default is the first sheet)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Ensure the first worksheet is the active one
        workbook.Worksheets.ActiveSheetIndex = 0;

        // Configure CSV save options: use semicolon as separator and export only the active sheet
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.Separator = ';';
        csvOptions.ExportAllSheets = false; // default, but set explicitly for clarity

        // Export the active sheet to a CSV file
        workbook.Save("ActiveSheet.csv", csvOptions);
    }
}