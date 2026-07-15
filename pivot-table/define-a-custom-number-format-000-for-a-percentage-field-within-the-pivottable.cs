using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotNumberFormatExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(0.1234); // 12.34%
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(0.5678); // 56.78%
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(0.9012); // 90.12%

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the category field to the row area
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the value field to the data area
            int dataFieldPos = pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            PivotField dataField = pivot.DataFields[dataFieldPos];

            // Set the custom number format to display as percentage with two decimal places
            dataField.NumberFormat = "0.00%";

            // Refresh and calculate the pivot table to apply the format
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableWithCustomNumberFormat.xlsx");
        }
    }
}