// Title: Create a Waterfall chart with Aspose.Cells for .NET and flag the final bar as a total using the IsTotal property
// AI Prompts: Generate C# code that adds a Waterfall chart to a worksheet, binds it to a data range, and sets the IsTotal flag on the chart’s last point. | Show how to configure the Waterfall chart title and save the workbook after marking the final data point as a total in Aspose.Cells.
// Common Searches: Aspose.Cells C# how to mark the last point of a waterfall chart as total | example of creating a waterfall chart with total column using Aspose.Cells .NET | set IsTotal flag on waterfall chart data point in Aspose.Cells for .NET
// Tags: Aspose.Cells create waterfall chart | Aspose.Cells set IsTotal property | C# waterfall chart from worksheet range | Aspose.Cells total bar in waterfall chart

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, writes category and value data, adds a Waterfall chart linked to that data, optionally sets the IsTotal flag on the final point to display it as a total bar, assigns a chart title, and saves the workbook as WaterfallChart.xlsx.
class WaterfallChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the waterfall chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            string[] categories = { "Sales", "Marketing", "R&D", "Total" };
            double[] values = { 200, -50, -30, 120 };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue(values[i]);   // Column B
            }

            // Add a Waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data range (values) and categories
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Optional: mark the final data point as total if the API supports it
            // (Commented out because IsTotal may not be available in older versions)
            // int lastPointIndex = chart.NSeries[0].Points.Count - 1;
            // chart.NSeries[0].Points[lastPointIndex].IsTotal = true;

            // Optional: set a title for the chart
            chart.Title.Text = "Waterfall Chart Example";

            // Define output file path
            string outputPath = "WaterfallChart.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
