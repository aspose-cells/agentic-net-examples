using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

class PivotChartWithColors
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a simple workbook with sample data
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Product");
                ws.Cells["B1"].PutValue("Quantity");
                ws.Cells["C1"].PutValue("Revenue");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["C2"].PutValue(100);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(20);
                ws.Cells["C3"].PutValue(200);
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B4"].PutValue(30);
                ws.Cells["C4"].PutValue(300);
                ws.Cells["A5"].PutValue("D");
                ws.Cells["B5"].PutValue(40);
                ws.Cells["C5"].PutValue(400);
                ws.Cells["A6"].PutValue("E");
                ws.Cells["B6"].PutValue(50);
                ws.Cells["C6"].PutValue(500);
                workbook.Save(inputPath, SaveFormat.Excel97To2003);
            }

            // Access the first worksheet (assumed to contain source data)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source data range for the pivot table
            string sourceRange = "A1:C6";

            // Add a new PivotTable at cell E3 with the name "SalesPivot"
            int pivotIndex = sheet.PivotTables.Add(sourceRange, "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the PivotTable fields
            pivot.AddFieldToArea(PivotFieldType.Row, 0);          // Product
            pivot.AddFieldToArea(PivotFieldType.Data, 1);         // Quantity
            pivot.AddFieldToArea(PivotFieldType.Data, 2);         // Revenue

            // Refresh and calculate pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a column chart that will become a PivotChart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Link the chart to the pivot table using the correct pivot name
            chart.PivotSource = pivot.Name;
            chart.RefreshPivotData();

            // Assign colors to each data series if they exist
            if (chart.NSeries.Count > 0 && chart.NSeries[0] != null)
                chart.NSeries[0].Area.ForegroundColor = Color.Green;
            if (chart.NSeries.Count > 1 && chart.NSeries[1] != null)
                chart.NSeries[1].Area.ForegroundColor = Color.Orange;

            // Optional: customize pivot controls on the chart
            PivotOptions pivOpts = chart.PivotOptions;
            pivOpts.DropZonesVisible = true;
            pivOpts.DropZoneData = true;
            pivOpts.DropZoneCategories = true;
            pivOpts.DropZoneSeries = true;
            pivOpts.DropZoneFilter = true;

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}