using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUIMinimal
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Item");
            sheet.Cells["C1"].PutValue("Amount");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Orange");
            sheet.Cells["C3"].PutValue(80);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(50);

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2); // Amount

            // Disable UI elements related to the pivot table
            pivotTable.EnableWizard = false;        // Hide the PivotTable wizard
            pivotTable.EnableFieldList = false;    // Hide the field list on the worksheet
            workbook.Settings.HidePivotFieldList = true; // Hide the global pivot field list ribbon

            // Save the workbook as ODS with default options
            workbook.Save("PivotTable_MinimalUI.ods", SaveFormat.Ods);
        }
    }
}