// Title: Generate a stacked column PivotChart with three data fields and custom series colors in an XLS workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing XLS file, creates a pivot table with three summed value fields, adds a stacked‑column PivotChart on a separate sheet, and sets a distinct RGB color for each series using Aspose.Cells. | Modify the example to produce a line PivotChart instead of a stacked column chart and change the series colors to custom shades of orange, teal, and purple.
// Common Searches: how to build a pivot chart with several data fields in Aspose.Cells using C# | setting individual series colors on a stacked column chart with Aspose.Cells | programmatically generate a pivot table and chart from an XLS workbook in .NET | custom series color assignment for Aspose.Cells charts | saving a workbook that contains a pivot chart to XLS format
// Tags: create pivot chart Aspose.Cells C# | stacked column chart series color customization | pivot table multiple data fields Aspose.Cells | assign series foreground color Aspose.Cells | save workbook with pivot chart XLS | programmatic chart generation Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// The sample ensures an input.xls file exists (creating sample data if missing), loads it, builds a pivot table with three summed value fields, adds a stacked column PivotChart on a new sheet, assigns distinct RGB colors to each series, sets a chart title, and saves the workbook as output.xls.
class PivotChartExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xls";

            // Ensure the input file exists; create a simple workbook if missing
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Data";

                // Populate sample data (A1:D5)
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value1");
                ws.Cells["C1"].PutValue("Value2");
                ws.Cells["D1"].PutValue("Value3");

                ws.Cells["A2"].PutValue("A");
                ws.Cells["A3"].PutValue("B");
                ws.Cells["A4"].PutValue("C");
                ws.Cells["A5"].PutValue("D");

                ws.Cells["B2"].PutValue(10);
                ws.Cells["B3"].PutValue(20);
                ws.Cells["B4"].PutValue(30);
                ws.Cells["B5"].PutValue(40);

                ws.Cells["C2"].PutValue(15);
                ws.Cells["C3"].PutValue(25);
                ws.Cells["C4"].PutValue(35);
                ws.Cells["C5"].PutValue(45);

                ws.Cells["D2"].PutValue(5);
                ws.Cells["D3"].PutValue(10);
                ws.Cells["D4"].PutValue(15);
                ws.Cells["D5"].PutValue(20);

                wb.Save(inputPath);
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Worksheet that contains source data (assumed first sheet)
            Worksheet dataSheet = workbook.Worksheets[0];

            // Define source data range (e.g., A1:D100)
            int firstRow = 0;          // zero‑based
            int firstColumn = 0;
            int totalRows = 100;
            int totalColumns = 4;
            string startCell = $"{CellsHelper.ColumnIndexToName(firstColumn)}{firstRow + 1}";
            string endCell = $"{CellsHelper.ColumnIndexToName(firstColumn + totalColumns - 1)}{firstRow + totalRows}";
            string sourceData = $"='{dataSheet.Name}'!{startCell}:{endCell}";

            // Add a worksheet to host the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Add the pivot table at cell A1 of the pivot sheet
            int pivotTableIndex = pivotSheet.PivotTables.Add("A1", sourceData, "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Add fields to the pivot table by index (0‑based)
            // Row field: Category (first column)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            // Data fields: Value1, Value2, Value3 (columns 1‑3)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 3);

            // Set aggregation type to Sum for each data field
            foreach (PivotField df in pivotTable.DataFields)
            {
                df.Function = ConsolidationFunction.Sum;
            }

            // Refresh pivot cache and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a worksheet for the chart
            int chartSheetIndex = workbook.Worksheets.Add();
            Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
            chartSheet.Name = "PivotChart";

            // Add a stacked column chart anchored at B2 (row 1, column 1) with size 20x15 cells
            int chartIndex = chartSheet.Charts.Add(ChartType.ColumnStacked, 1, 1, 20, 15);
            Chart chart = chartSheet.Charts[chartIndex];

            // Determine the data range for the chart (including header row)
            int dataRows = dataSheet.Cells.MaxDataRow + 1; // total rows in source data (header + data)
            string chartStart = "A1";
            string chartEnd = $"{CellsHelper.ColumnIndexToName(totalColumns - 1)}{dataRows}";
            string chartDataRange = $"='Pivot'!{chartStart}:{chartEnd}";

            // Set the chart's data source
            chart.NSeries.Add(chartDataRange, true);

            // Assign colors to each series (if series exist)
            if (chart.NSeries.Count > 0)
                chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(255, 79, 129, 189);   // Blue
            if (chart.NSeries.Count > 1)
                chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(255, 192, 80, 77);   // Red
            if (chart.NSeries.Count > 2)
                chart.NSeries[2].Area.ForegroundColor = Color.FromArgb(255, 155, 187, 89);  // Green

            // Optional chart title
            chart.Title.Text = "Pivot Chart with Multiple Value Fields";

            // Save the workbook with the new pivot chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
