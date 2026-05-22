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
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or specify the required one)
            Worksheet worksheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 1. Create a PivotTable from range A1:E20
            // -----------------------------------------------------------------
            // Add the PivotTable at cell G3 with the name "PivotTable1"
            int pivotIndex = worksheet.PivotTables.Add("A1:E20", "G3", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the PivotTable (example: first column as Row field, second as Data field)
            // Adjust field indices or names according to your source data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // First column as Row
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Second column as Data

            // Refresh and calculate the PivotTable data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -----------------------------------------------------------------
            // 2. Add a linked PivotChart (PivotChart) to the same worksheet
            // -----------------------------------------------------------------
            // Create a column chart positioned at rows 15-30 and columns 0-7
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 15, 0, 30, 7);
            Chart chart = worksheet.Charts[chartIndex];

            // Link the chart to the PivotTable created above
            // Since the chart and pivot table are in the same workbook, use the simple reference
            chart.PivotSource = $"{worksheet.Name}!{pivotTable.Name}";

            // Refresh the chart to reflect the PivotTable data
            chart.RefreshPivotData();

            // -----------------------------------------------------------------
            // 3. Save the modified workbook
            // -----------------------------------------------------------------
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}