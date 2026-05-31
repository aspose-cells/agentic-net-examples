using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerChartDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // 1. Populate sample data (Fruit, Sales)
                // -------------------------------------------------
                cells["A1"].PutValue("Fruit");
                cells["B1"].PutValue("Sales");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(120);
                cells["A3"].PutValue("Orange");
                cells["B3"].PutValue(150);
                cells["A4"].PutValue("Banana");
                cells["B4"].PutValue(90);
                cells["A5"].PutValue("Grape");
                cells["B5"].PutValue(60);

                // -------------------------------------------------
                // 2. Create a PivotTable based on the data
                // -------------------------------------------------
                // Place the pivot table starting at cell D1
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];

                // Row field: Fruit, Data field: Sales (Sum)
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh to calculate the pivot data
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // 3. Add a chart that is linked to the PivotTable
                // -------------------------------------------------
                // The chart will be placed at rows 12‑22, columns 0‑7
                int chartIdx = sheet.Charts.Add(ChartType.Column, 12, 0, 22, 7);
                Chart chart = sheet.Charts[chartIdx];
                chart.Title.Text = "Sales by Fruit (Pivot)";

                // Use a static address that covers the pivot table data area.
                // After the pivot is calculated, the data starts at D2 and ends at E5.
                chart.NSeries.Add("D2:E5", true);
                chart.NSeries[0].Name = "Sales";

                // -------------------------------------------------
                // 4. Add a slicer linked to the PivotTable's "Fruit" field
                // -------------------------------------------------
                // Place the slicer at cell G1 (row 0, column 6)
                int slicerIdx = sheet.Slicers.Add(pivot, 0, 6, "Fruit");
                Slicer slicer = sheet.Slicers[slicerIdx];
                slicer.Caption = "Fruit Filter";
                slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                workbook.Save("SlicerChartDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}