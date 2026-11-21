// Title: Generate a Pie Chart from cells A1:A5 and place it on a new worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create pie chart from a range and add it to a separate worksheet | How to add a pie chart to a new sheet in an Excel file with Aspose.Cells .NET | C# example for inserting a pie chart using Aspose.Cells and setting data source to Data!A1:A5 | Aspose.Cells generate chart on another worksheet programmatically | Create chart sheet with pie chart in Aspose.Cells C# tutorial
// Tags: pie chart generation Aspose.Cells C# | add chart to new worksheet Aspose.Cells | chart data source range Aspose.Cells | Aspose.Cells chart sheet creation .NET | Excel pie chart programmatic insertion C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart and ChartType

// The sample creates a new workbook, fills cells A1‑A5 on a worksheet named "Data" with numeric values, adds a second worksheet called "ChartSheet", inserts a Pie chart positioned at the top‑left of the new sheet, links the chart's series to the range Data!A1:A5, sets a title, and saves the file as "PieChartExample.xlsx" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (default) and fill sample data in A1:A5
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";
            dataSheet.Cells["A1"].PutValue(10);
            dataSheet.Cells["A2"].PutValue(20);
            dataSheet.Cells["A3"].PutValue(30);
            dataSheet.Cells["A4"].PutValue(25);
            dataSheet.Cells["A5"].PutValue(15);

            // Add a new worksheet to host the chart
            Worksheet chartSheet = workbook.Worksheets.Add("ChartSheet");

            // Add a pie chart to the new worksheet (positioned at row 0, column 0, spanning 20 rows and 10 columns)
            int chartIndex = chartSheet.Charts.Add(ChartType.Pie, 0, 0, 20, 10);
            Chart pieChart = chartSheet.Charts[chartIndex];

            // Set the data source for the pie chart (range A1:A5 on the Data sheet)
            pieChart.NSeries.Add("Data!A1:A5", true);

            // Optional: set chart title
            pieChart.Title.Text = "Sample Pie Chart";

            // Save the workbook to a file
            workbook.Save("PieChartExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
