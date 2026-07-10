using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Ods;

namespace PivotRibbonDisableDemo
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
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(850);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(430);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product column as rows
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column as values

            // Configure ODS save options to ignore pivot tables (removes pivot ribbons)
            OdsSaveOptions saveOptions = new OdsSaveOptions
            {
                IgnorePivotTables = true   // Disables pivot table ribbons in the ODS output
            };

            // Save the workbook as ODS with the specified options
            workbook.Save("CleanPivotTable.ods", saveOptions);
        }
    }
}