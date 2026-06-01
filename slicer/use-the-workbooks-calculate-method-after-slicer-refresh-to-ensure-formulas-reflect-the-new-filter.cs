using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCalculateDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet (data source)
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(150);

            // Add a worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");
            int pivotIdx = pivotSheet.PivotTables.Add("A1:B4", "C3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIdx];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales

            // Add a worksheet that will host the slicer
            Worksheet slicerSheet = workbook.Worksheets.Add("SlicerSheet");
            int slicerIdx = slicerSheet.Slicers.Add(pivotTable, "A1", "Product");
            Slicer slicer = slicerSheet.Slicers[slicerIdx];

            // Insert a formula that depends on the pivot table (GETPIVOTDATA example)
            // This formula will sum sales for "Apple". It will be recalculated after the slicer refresh.
            pivotSheet.Cells["E2"].Formula = "=GETPIVOTDATA(\"Sales\",C3,\"Product\",\"Apple\")";

            // Change source data to demonstrate that slicer filtering affects the pivot and formula
            dataSheet.Cells["A2"].PutValue("Apple");   // keep Apple
            dataSheet.Cells["A3"].PutValue("Apple");   // change Banana to Apple
            dataSheet.Cells["B3"].PutValue(200);       // new sales value for the added Apple row

            // Refresh the slicer – this also refreshes the underlying pivot table
            slicer.Refresh();

            // After slicer refresh, recalculate all formulas so that the GETPIVOTDATA result is up‑to‑date
            workbook.CalculateFormula();

            // Output the calculated result to the console
            Console.WriteLine("Calculated sales for Apple (E2): " + pivotSheet.Cells["E2"].StringValue);

            // Save the workbook (using the standard save rule)
            workbook.Save("SlicerRefreshWithCalculate.xlsx");
        }
    }
}