using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Ods;

namespace AsposeCellsPivotOdsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(3000);

            // Create a pivot table based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:B4", "E5", "PivotTable1");
            PivotTable pivot = pivots[pivotIndex];

            // Add fields: Product as row, Sales as data
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Product
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column index 1 -> Sales

            // Configure ODS save options
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            // Include pivot tables in the saved ODS file
            saveOptions.IgnorePivotTables = false;
            // Optional: specify the generator type (LibreOffice or OpenOffice)
            saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

            // Save the workbook as an ODS file with the specified options
            workbook.Save("PivotTableDemo.ods", saveOptions);
        }
    }
}