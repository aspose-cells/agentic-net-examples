// Title: How to set a 2‑D clustered bar chart gap width to 150 % using Aspose.Cells for .NET
// AI Prompts: Provide C# code that creates a workbook, adds sample data, inserts a clustered bar chart, and sets its GapWidth property to 150 percent with Aspose.Cells. | Show an example of tightening column spacing in an Excel bar chart by configuring the GapWidth value to 150 in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set bar chart gap width to 150 percent | change column spacing of clustered bar chart in .NET using Aspose.Cells | adjust gap width property for Excel bar chart programmatically | tighten bar chart columns Aspose.Cells example | modify bar chart gap width in a generated XLSX file with C#
// Tags: Aspose.Cells GapWidth property for bar charts | C# set bar chart column spacing | clustered bar chart gap width adjustment | Excel chart formatting with Aspose.Cells | generate bar chart with custom spacing .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills it with sample data, adds a 2‑D clustered bar chart, assigns the series and categories, sets the chart's GapWidth to 150 % for tighter column spacing, and saves the file as BarChartWithAdjustedGapWidth.xlsx.
    public class AdjustBarChartGapWidth
    {
        // Entry point for the example
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the bar chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a 2‑D bar chart (clustered)
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Adjust the gap width to 150 % (tighter column spacing)
            chart.GapWidth = 150;   // Valid range: 0‑500

            // Save the workbook to a file
            string outputPath = "BarChartWithAdjustedGapWidth.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
