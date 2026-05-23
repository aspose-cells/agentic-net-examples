using System;
using Aspose.Cells;
using System.Text;

class SaveWorkbookAsCsv
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate cells with mixed data types
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Date");
        cells["C1"].PutValue("Amount");

        cells["A2"].PutValue("John");
        cells["B2"].PutValue(DateTime.Now);          // DateTime value
        cells["C2"].PutValue(123.45);                // Numeric value

        cells["A3"].PutValue("Alice");
        cells["B3"].PutValue(new DateTime(2023, 5, 1));
        cells["C3"].PutValue(987);                   // Integer value

        // Ensure cells keep their original data types (do not force all to strings)
        cells.PreserveString = false;

        // Configure CSV save options
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.Separator = ',';          // Use comma as delimiter
        csvOptions.ClearData = false;        // Keep workbook data after saving
        csvOptions.ExportAllSheets = false;  // Export only the active sheet

        // Save the workbook as CSV while preserving data types
        workbook.Save("output.csv", csvOptions);
    }
}