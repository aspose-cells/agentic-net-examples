using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsPivotChartDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook --------------------
            Workbook workbook = new Workbook(); // create a new workbook

            // -------------------- Add source data worksheet --------------------
            Worksheet sourceSheet = workbook.Worksheets[0]; // first sheet will hold the source data
            sourceSheet.Name = "SourceData";

            // Populate sample data
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Product");
            sourceSheet.Cells["C1"].PutValue("Sales");

            sourceSheet.Cells["A2"].PutValue("Fruit");
            sourceSheet.Cells["B2"].PutValue("Apple");
            sourceSheet.Cells["C2"].PutValue(1200);

            sourceSheet.Cells["A3"].PutValue("Fruit");
            sourceSheet.Cells["B3"].PutValue("Banana");
            sourceSheet.Cells["C3"].PutValue(850);

            sourceSheet.Cells["A4"].PutValue("Vegetable");
            sourceSheet.Cells["B4"].PutValue("Carrot");
            sourceSheet.Cells["C4"].PutValue(640);

            sourceSheet.Cells["A5"].PutValue("Vegetable");
            sourceSheet.Cells["B5"].PutValue("Tomato");
            sourceSheet.Cells["C5"].PutValue(970);

            // -------------------- Add pivot table worksheet --------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Define source data range in A1 style (including sheet name)
            string sourceData = $"=SourceData!{sourceSheet.Cells.MaxDisplayRange.Address}";

            // Add pivot table at cell A1 of the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // rows: Category
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product"); // columns: Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // values: Sales

            // Optional: format layout
            pivotTable.ShowInTabularForm();
            pivotTable.CalculateData(); // populate the pivot table

            // -------------------- Add chart worksheet --------------------
            Worksheet chartSheet = workbook.Worksheets.Add("PivotChart");

            // Add a column chart
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
            Chart chart = chartSheet.Charts[chartIndex];

            // Use the source data range for the chart series
            string chartSource = $"SourceData!{sourceSheet.Cells.MaxDisplayRange.Address}";
            chart.NSeries.Add(chartSource, true);

            // Set chart title (optional)
            chart.Title.Text = "Sales by Category and Product";

            // -------------------- Save the workbook --------------------
            workbook.Save("PivotTableWithLinkedChart.xlsx");
        }
    }
}