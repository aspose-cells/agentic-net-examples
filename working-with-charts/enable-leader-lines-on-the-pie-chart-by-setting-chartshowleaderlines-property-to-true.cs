// Title: Add leader lines to a pie chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, inserts a pie chart, and enables leader lines on its series with Aspose.Cells. | Show how to configure the HasLeaderLines property for an NSeries object in Aspose.Cells. | Provide a complete example that saves an Excel file containing a pie chart with visible leader lines.
// Common Searches: how to enable leader lines on a pie chart with Aspose.Cells C# | Aspose.Cells C# example for adding leader lines to Excel pie chart series | set HasLeaderLines true for pie chart series using Aspose.Cells .NET | C# code to display leader lines in Excel pie chart via Aspose.Cells
// Tags: Aspose.Cells pie chart leader lines | C# HasLeaderLines property Aspose.Cells | NSeries visual formatting Aspose.Cells | Excel pie chart leader line generation C# | Aspose.Cells chart series customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLeaderLinesDemo
{
    // The sample creates a new workbook, fills it with category and value data, adds a pie chart, assigns the data range, turns on leader lines for the chart series by setting HasLeaderLines to true, and saves the result as PieChartWithLeaderLines.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a pie chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable leader lines for the first series (pie chart)
            chart.NSeries[0].HasLeaderLines = true;

            // Save the workbook to a file
            workbook.Save("PieChartWithLeaderLines.xlsx");
        }
    }
}
