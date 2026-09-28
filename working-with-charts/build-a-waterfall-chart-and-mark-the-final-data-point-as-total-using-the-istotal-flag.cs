// Title: Generate a Waterfall Chart in Excel with Aspose.Cells for .NET and mark the last point as a total using the IsTotal flag
// AI Prompts: Create C# code that uses Aspose.Cells to populate a worksheet with category and value data, adds a Waterfall chart, and sets the final data point's IsTotal property to true. | Show how to define the series range and category range for a Waterfall chart in Aspose.Cells, customize the chart title, and position the chart within the worksheet. | Demonstrate saving the workbook that contains the Waterfall chart to an .xlsx file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells .NET how to set IsTotal on a waterfall chart data point | C# create waterfall chart in Excel and mark total column using Aspose.Cells | example of adding a waterfall chart with total bar in Aspose.Cells workbook | set final bar as total in Aspose.Cells waterfall chart programmatically
// Tags: Aspose.Cells waterfall chart creation | C# IsTotal flag usage | Excel worksheet data preparation for chart | Aspose.Cells chart positioning | save workbook as .xlsx with chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Creates a new workbook, fills columns A and B with categories and values, adds a Waterfall chart, assigns series and category ranges, marks the final point as total via IsTotal, sets a chart title, positions the chart, and saves the file as WaterfallChart.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate data for the waterfall chart
            // Column A: Categories, Column B: Values
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Start");
            ws.Cells["B2"].PutValue(100);
            ws.Cells["A3"].PutValue("Increase");
            ws.Cells["B3"].PutValue(30);
            ws.Cells["A4"].PutValue("Decrease");
            ws.Cells["B4"].PutValue(-20);
            ws.Cells["A5"].PutValue("Total");
            ws.Cells["B5"].PutValue(110); // Final total value

            // Add a waterfall chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = ws.Charts.Add(ChartType.Waterfall, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Optional: set chart title
            chart.Title.Text = "Waterfall Chart Example";

            // Define output file path
            string outputPath = "WaterfallChart.xlsx";

            // Save the workbook to a file
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
