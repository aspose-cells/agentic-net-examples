// Title: Adjust Bar Chart GapWidth to 150 % (tighter columns) with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add sample data, insert a column chart, set its GapWidth property to 150 % (default = 100) to tighten bar spacing, and save the file as XLSX using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# chart GapWidth | adjust column spacing | Excel bar chart gap width | Chart.GapWidth property | Aspose.Cells .NET | column chart spacing | set GapWidth 150 | tight bar chart | Excel chart formatting
// Common Searches: Aspose.Cells set chart gap width C# | How to change column chart spacing in Aspose.Cells | GapWidth property example Aspose.Cells .NET | Reduce space between bars Aspose.Cells | Adjust bar chart gap width percentage
// Developer Intent: The developer wants to set the GapWidth of a column (bar) chart to 150 % so the columns appear closer together.
// Use Cases: Generate a sales dashboard where column charts use a compact layout to display many categories. | Create financial reports that require tightly spaced bars to fit more data points on a single chart. | Automate batch creation of Excel charts with varying GapWidth values for visual density comparison.
// AI Prompts: Show how to set Chart.GapWidth to 150 in Aspose.Cells for .NET. | Provide a C# snippet that adjusts a column chart's gap width based on a percentage variable. | Explain the valid range for the GapWidth property and its visual impact on bar charts in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add sample data, insert a column chart, set its GapWidth property to 150 % (default = 100) to tighten bar spacing, and save the file as XLSX using Aspose.Cells for C#.
    public class AdjustBarChartGapWidth
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column (bar) chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Adjust the gap width to 150% (tighter column spacing)
                // Valid range is 0 to 500; 150 means 150% of the default width
                chart.GapWidth = 150;

                // Save the workbook to a file
                workbook.Save("AdjustedGapWidthChart.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully as AdjustedGapWidthChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustBarChartGapWidth.Run();
        }
    }
}
