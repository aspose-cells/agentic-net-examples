using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load an existing workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets.Clear(); // start with a clean sheet collection
                workbook.Worksheets.Add("Sheet1");
            }

            // Use the first worksheet
            Worksheet ws = workbook.Worksheets[0];

            // Prepare sample data for the pivot table (overwrites any existing data)
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("A");
            ws.Cells["B4"].PutValue(30);
            ws.Cells["A5"].PutValue("B");
            ws.Cells["B5"].PutValue(40);

            // Add a pivot table based on the data range
            int pivotIdx = ws.PivotTables.Add("=A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = ws.PivotTables[pivotIdx];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a chart that will become a PivotChart
            int chartIdx = ws.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];
            chart.PivotSource = "PivotTable1"; // Link chart to the pivot table

            // Refresh chart data from the pivot table
            chart.RefreshPivotData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}