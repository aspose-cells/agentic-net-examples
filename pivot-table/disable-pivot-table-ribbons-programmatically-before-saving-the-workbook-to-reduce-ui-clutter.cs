using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRibbonDisable
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(2000);

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D5", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales column

            // Disable UI elements related to pivot tables
            // Hide the PivotTable field list ribbon (reduces UI clutter)
            workbook.Settings.HidePivotFieldList = true;

            // Additionally, turn off the PivotTable wizard and field list for this specific pivot table
            pivotTable.EnableWizard = false;
            pivotTable.EnableFieldList = false;

            // Save the workbook (saving rule)
            workbook.Save("PivotTable_RibbonDisabled.xlsx");
        }
    }
}