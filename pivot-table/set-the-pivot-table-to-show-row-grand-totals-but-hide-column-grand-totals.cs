using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (will hold the source data)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            // Columns: Product, Region, Sales
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Region");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Product A");
            dataSheet.Cells["B2"].PutValue("North");
            dataSheet.Cells["C2"].PutValue(1000);

            dataSheet.Cells["A3"].PutValue("Product B");
            dataSheet.Cells["B3"].PutValue("South");
            dataSheet.Cells["C3"].PutValue(1500);

            dataSheet.Cells["A4"].PutValue("Product A");
            dataSheet.Cells["B4"].PutValue("South");
            dataSheet.Cells["C4"].PutValue(2000);

            dataSheet.Cells["A5"].PutValue("Product B");
            dataSheet.Cells["B5"].PutValue("North");
            dataSheet.Cells["C5"].PutValue(1200);

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table using the source range A1:C5 and place it at E3
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            // Row: Product (field index 0)
            // Column: Region (field index 1)
            // Data: Sales (field index 2)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);

            // Set grand total visibility
            pivotTable.ShowRowGrandTotals = true;   // Show row grand totals
            pivotTable.ShowColumnGrandTotals = false; // Hide column grand totals

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_RowGrand_HideColumnGrand.xlsx");
        }
    }
}