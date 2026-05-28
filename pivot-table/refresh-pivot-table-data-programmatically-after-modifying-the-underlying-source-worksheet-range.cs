using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace RefreshPivotExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(1200);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(850);
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(950);

            // Add a pivot table based on the source data
            int pivotIndex = dataSheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Initial calculation so the pivot shows the original data
            pivotTable.CalculateData();

            // ----- Modify the underlying source data -----
            dataSheet.Cells["B2"].PutValue(1300); // Apple sales changed
            dataSheet.Cells["B3"].PutValue(900);  // Banana sales changed

            // Refresh all pivot tables in the worksheet to reflect the changes
            dataSheet.RefreshPivotTables();

            // Optionally, recalculate the pivot data after refresh (not required for RefreshPivotTables)
            // pivotTable.CalculateData();

            // Save the workbook with refreshed pivot data
            workbook.Save("RefreshedPivot.xlsx");
        }
    }
}