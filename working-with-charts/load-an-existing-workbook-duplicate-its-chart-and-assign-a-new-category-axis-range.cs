// Title: How to duplicate a chart in an Excel workbook and set a new category axis range with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a workbook, clones the first chart on a worksheet, copies all its series, and assigns a different CategoryData range to the cloned chart using Aspose.Cells. | Write a method that creates a sample workbook if it does not exist, adds a copy of an existing chart with the same type and position, and updates the copied chart's category axis to a custom cell range.
// Common Searches: Aspose.Cells C# duplicate chart and change category axis range | Copy Excel chart programmatically and set new CategoryData using Aspose.Cells | How to clone a chart and modify its series axis in a .NET workbook with Aspose.Cells
// Tags: clone chart Aspose.Cells C# | set CategoryData range Aspose.Cells | copy chart series Aspose.Cells | add chart with same type Aspose.Cells | modify chart axis range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook (or creates a sample one), clones the first chart preserving its type and position, copies each series to the new chart, assigns a new category axis range (e.g., A2:A5), and saves the workbook with the duplicated chart.
class DuplicateChartExample
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "SourceWorkbook.xlsx";
            string outputPath = "WorkbookWithDuplicatedChart.xlsx";

            // Ensure the source workbook exists; create a simple one if missing
            if (!File.Exists(inputPath))
            {
                CreateSampleWorkbook(inputPath);
                Console.WriteLine($"Sample workbook created at '{inputPath}'.");
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (assumed to contain the chart)
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that at least one chart exists
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No chart found in the worksheet.");
                return;
            }

            // Get the first chart as the source chart
            Chart sourceChart = sheet.Charts[0];

            // Add a new chart with the same type and position as the source chart
            int newChartIndex = sheet.Charts.Add(
                sourceChart.Type,
                sourceChart.ChartObject.UpperLeftRow,
                sourceChart.ChartObject.UpperLeftColumn,
                sourceChart.ChartObject.LowerRightRow,
                sourceChart.ChartObject.LowerRightColumn);

            Chart newChart = sheet.Charts[newChartIndex];

            // Copy each series from the source chart to the new chart
            foreach (Series sourceSeries in sourceChart.NSeries)
            {
                // Add the series values to the new chart (true = isVertical)
                int seriesIdx = newChart.NSeries.Add(sourceSeries.Values, true);
                Series newSeries = newChart.NSeries[seriesIdx];

                // Copy series name (and other properties as needed)
                newSeries.Name = sourceSeries.Name;
            }

            // Assign a new category axis range to the duplicated chart (example: A2:A5)
            newChart.NSeries.CategoryData = "A2:A5";

            // Save the workbook with the duplicated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Chart duplicated and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to create a simple workbook with sample data and a chart
    private static void CreateSampleWorkbook(string path)
    {
        try
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["A5"].PutValue("D");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);
            ws.Cells["B5"].PutValue(40);

            // Add a column chart
            int chartIndex = ws.Charts.Add(ChartType.Column, 7, 0, 20, 7);
            Chart chart = ws.Charts[chartIndex];
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            wb.Save(path);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
        }
    }
}
