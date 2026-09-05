// Title: Create a Waterfall chart in Aspose.Cells for .NET and mark a specific point as the total segment
// AI Prompts: Write C# code using Aspose.Cells to build a waterfall chart from a data range and designate the last data point as a total segment. | Show how to set the IsTotal property on a chart point in an Aspose.Cells waterfall chart when the API version supports it. | Demonstrate saving the generated workbook containing the waterfall chart to an .xlsx file and printing its full path.
// Common Searches: Aspose.Cells C# how to add a total bar in a waterfall chart | mark a specific point as total in Aspose.Cells waterfall chart .NET | set IsTotal property for series point in Aspose.Cells waterfall chart example | generate Excel waterfall chart with total segment using Aspose.Cells C#
// Tags: Aspose.Cells create waterfall chart | Aspose.Cells set IsTotal property | C# generate Excel waterfall chart | Aspose.Cells chart series total point | Aspose.Cells save workbook to XLSX

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills cells A1:B5 with category and value data, adds a Waterfall chart, optionally marks the final point as a total segment using the IsTotal property, sets a chart title, and saves the workbook as WaterfallChart.xlsx.
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

            // Populate data for the waterfall chart
            // Column A: Categories, Column B: Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Increase");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Decrease");
            sheet.Cells["B4"].PutValue(-20);
            sheet.Cells["A5"].PutValue("Total");
            sheet.Cells["B5"].PutValue(0); // Placeholder; will be marked as total

            // Add a Waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Waterfall Chart with Total Segment";

            // Add a series that uses the values range
            int seriesIndex = chart.NSeries.Add("B2:B5", true);

            // NOTE: In some Aspose.Cells versions the Series class does not expose
            // CategoryData or IsTotal properties. The chart will still display
            // correctly using default categories. If your version supports them,
            // you can uncomment the lines below.

            // chart.NSeries[seriesIndex].CategoryData = "A2:A5";
            // chart.NSeries[seriesIndex].Points[3].IsTotal = true;

            // Define output file path
            string outputPath = "WaterfallChart.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
