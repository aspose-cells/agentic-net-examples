using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRefreshCalculateDemo
{
    public class Program
    {
        public static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate source data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(150);

            // ---------- Create a pivot table based on the data ----------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIdx = pivotSheet.PivotTables.Add("A1:B4", "C3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIdx];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales

            // ---------- Add a slicer linked to the pivot table ----------
            Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
            int slicerIdx = slicerSheet.Slicers.Add(pivotTable, "A1", "Product");
            Slicer slicer = slicerSheet.Slicers[slicerIdx];

            // ---------- Add a formula that reads the total sales from the pivot ----------
            // Using GETPIVOTDATA to fetch the sum of Sales for all products
            Worksheet resultSheet = workbook.Worksheets.Add("Result");
            resultSheet.Cells["A1"].PutValue("Total Sales (Pivot)");
            resultSheet.Cells["B1"].Formula = "=GETPIVOTDATA(\"Sales\",\"Pivot!C3\")";

            // ---------- Change source data to simulate user interaction ----------
            dataSheet.Cells["B2"].PutValue(200); // Apple sales changed
            dataSheet.Cells["B3"].PutValue(90);  // Banana sales changed

            // ---------- Refresh the slicer (which also refreshes the pivot) ----------
            slicer.Refresh();

            // ---------- Recalculate formulas so that GETPIVOTDATA reflects the refreshed pivot ----------
            workbook.CalculateFormula();

            // ---------- Save the workbook ----------
            workbook.Save("SlicerRefreshWithCalculate.xlsx");
        }
    }
}