using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // -------------------- Data Worksheet --------------------
            Worksheet dataSheet = wb.Worksheets[0];
            dataSheet.Name = "Data";

            // Add header cells and merge them (merged header demonstrates merged data handling)
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells.Merge(0, 0, 1, 2); // Merge A1:B1

            // Insert smart markers for the data rows (starting at row 2)
            // The smart markers will be replaced with values from the data source named "Items"
            dataSheet.Cells["A2"].PutValue("&=$Items.Category");
            dataSheet.Cells["B2"].PutValue("&=$Items.Value");

            // Define the smart marker range (required for processing)
            AsposeRange smRange = dataSheet.Cells.CreateRange("A2:B2");
            smRange.Name = "_CellsSmartMarkers";

            // -------------------- Chart Worksheet --------------------
            Worksheet chartSheet = wb.Worksheets.Add("Chart");

            // Add a column chart placeholder
            int chartIdx = chartSheet.Charts.Add(ChartType.Column, 2, 1, 20, 10);
            Chart chart = chartSheet.Charts[chartIdx];

            // Initially set a provisional data range; it will be updated after processing
            chart.SetChartDataRange("Data!$A$1:$B$5", true);

            // -------------------- Data Source --------------------
            var items = new List<Item>()
            {
                new Item { Category = "A", Value = 10 },
                new Item { Category = "B", Value = 20 },
                new Item { Category = "C", Value = 30 },
                new Item { Category = "D", Value = 40 }
            };

            // -------------------- Process Smart Markers --------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb
            };
            designer.SetDataSource("Items", items);
            designer.Process(); // Process all smart markers in the workbook

            // After processing, adjust the chart data range to cover the actual populated rows
            int lastDataRow = dataSheet.Cells.MaxDataRow; // zero‑based index
            // +1 because Excel rows are 1‑based and we want to include the header row
            string actualRange = $"Data!$A$1:$B${lastDataRow + 1}";
            chart.SetChartDataRange(actualRange, true);

            // -------------------- Save Result --------------------
            string outputPath = "SmartMarkerChart.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Simple POCO class used as the data source for smart markers
    public class Item
    {
        public string Category { get; set; } = string.Empty;
        public double Value { get; set; }
    }
}