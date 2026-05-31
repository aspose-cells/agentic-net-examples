using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(30);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(40);

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

            // Create the pivot table (source range, destination cell, name)
            int pivotIndex = pivotSheet.PivotTables.Add("Sheet1!A1:B5", "C3", "MyPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Disable automatic refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = false;

            // Optional: enable manual update so the pivot does not recalc automatically
            pivotTable.ManualUpdate = true;

            // Initial calculation to populate the pivot table
            pivotTable.CalculateData();

            // Save the workbook before any user edits
            workbook.Save("PivotBeforeEdit.xlsx");

            // -------------------------------------------------
            // Simulate user editing the source data after the file is opened
            // (In a real scenario the user would edit the file manually)
            // -------------------------------------------------
            dataSheet.Cells["B2"].PutValue(100); // Change value for Category A
            dataSheet.Cells["B3"].PutValue(200); // Change value for Category B

            // Because RefreshDataOnOpeningFile is false and ManualUpdate is true,
            // the pivot table will not refresh automatically.
            // Manually refresh the pivot cache and recalculate the pivot table.
            pivotTable.RefreshData();      // Gather data from the source range
            pivotTable.CalculateData();   // Recalculate the pivot results

            // Save the workbook after manual refresh
            workbook.Save("PivotAfterManualRefresh.xlsx");
        }
    }
}