using System;
using Aspose.Cells;

namespace AsposeCellsFodsDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook (default format is Xlsx)
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apples");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["A3"].PutValue("Bananas");
            worksheet.Cells["B3"].PutValue(85);

            // ------------------------------------------------------------
            // 2. Save the workbook as OpenDocument Flat XML Spreadsheet (FODS)
            // ------------------------------------------------------------
            workbook.Save("Sample.fods", SaveFormat.Fods);

            // ------------------------------------------------------------
            // 3. Load the previously saved FODS file
            // ------------------------------------------------------------
            Workbook loadedWorkbook = new Workbook("Sample.fods");

            // Access its first worksheet
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

            // Add a new row of data
            loadedWorksheet.Cells["A4"].PutValue("Cherries");
            loadedWorksheet.Cells["B4"].PutValue(60);

            // ------------------------------------------------------------
            // 4. Save the modified workbook back to FODS format
            // ------------------------------------------------------------
            loadedWorkbook.Save("Sample_Modified.fods", SaveFormat.Fods);
        }
    }
}