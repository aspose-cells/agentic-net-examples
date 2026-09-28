// Title: How to resize an Aspose.Cells chart to 500 pt × 300 pt and position it at cell D5 using C#
// AI Prompts: Set the WidthPt and HeightPt of a ChartObject to 500 and 300, then move the chart to the upper‑left corner of cell D5 with Aspose.Cells for .NET. | Create a column chart, assign sample data, resize it to 500 points wide by 300 points high, align its top‑left corner with D5, and save the workbook as an XLSX file.
// Common Searches: Aspose.Cells C# resize chart to specific point dimensions | move chart to cell D5 in Excel workbook using Aspose.Cells .NET | set chart width 500pt height 300pt programmatically with Aspose.Cells
// Tags: chartobject widthpt heightpt property Aspose.Cells | chart placement at specific cell Aspose.Cells | set chart dimensions points C# Aspose.Cells | column chart generation Aspose.Cells .NET | export workbook to xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartResize
{
    // The example creates a new workbook, adds a column chart with sample data, sets the chart's width to 500 pt and height to 300 pt, moves its upper‑left corner to cell D5, and saves the file as ResizedChart.xlsx.
    public class ResizeChartExample
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a sample column chart (initial position and size are temporary)
            // Parameters: ChartType, upper left row, upper left column, lower right row, lower right column
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // OPTIONAL: set some data for the chart so it is visible
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Resize the chart to 500 points width and 300 points height
            chart.ChartObject.WidthPt = 500;
            chart.ChartObject.HeightPt = 300;

            // Position the chart so its upper‑left corner aligns with cell D5 (row 4, column 3)
            chart.Move(4, 3, 4, 3);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ResizedChart.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ResizeChartExample.Run();
                Console.WriteLine("Workbook saved successfully as ResizedChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
