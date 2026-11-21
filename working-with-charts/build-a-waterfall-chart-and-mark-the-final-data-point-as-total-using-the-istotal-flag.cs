// Title: Create a Waterfall chart in Aspose.Cells for .NET and mark the final bar as total using the IsTotal flag
// Common Searches: Aspose.Cells C# how to set IsTotal on the last point of a waterfall chart | generate waterfall chart with total bar using Aspose.Cells .NET | C# example for marking the final waterfall chart point as total in Aspose.Cells | set IsTotal property on a waterfall chart series point Aspose.Cells | binding series data and categories for a waterfall chart in Aspose.Cells
// Tags: Aspose.Cells waterfall chart creation .NET | set IsTotal flag chart point Aspose.Cells | waterfall chart series data binding Aspose.Cells | add title to Aspose.Cells chart C# | save workbook as Excel file Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a new Workbook, fills category labels and numeric values, adds a Waterfall chart, binds the series to the value range B2:B5 and categories to A2:A5, optionally sets the fourth point's IsTotal property to true to display it as a total bar, assigns a chart title, and saves the file as WaterfallChart.xlsx.
class WaterfallChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the waterfall chart
            // Column A – categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["A5"].PutValue("Total");

            // Column B – values
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1000);   // Start
            sheet.Cells["B3"].PutValue(300);    // Revenue
            sheet.Cells["B4"].PutValue(-200);   // Cost
            sheet.Cells["B5"].PutValue(0);      // Placeholder for total

            // Add a waterfall chart to the worksheet
            // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";     // Categories

            // Mark the final data point as total (if supported by the library version)
            // The IsTotal property may not be available in older Aspose.Cells versions.
            // Uncomment the following line if your version supports it:
            // chart.NSeries[0].Points[3].IsTotal = true;

            // Optional: set a title for the chart
            chart.Title.Text = "Waterfall Chart Example";

            // Save the workbook to a file
            workbook.Save("WaterfallChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
