// Title: Create a new worksheet and insert a 400x300 pixel pie chart from range Data!A1:A5 using Aspose.Cells for .NET
// AI Prompts: Create a worksheet named "Chart" and add a Pie chart positioned at (0,0) with width 400 and height 300 pixels. | Link the Pie chart series to the data range "Data!A1:A5" on the source worksheet. | Set the chart title to "Sample Pie Chart" and save the workbook as an XLSX file.
// Common Searches: Aspose.Cells .NET how to add a pie chart on a separate sheet | reference data from another worksheet for a chart in Aspose.Cells C# | set chart dimensions in pixels using Aspose.Cells API | assign title to a pie chart created with Aspose.Cells for .NET
// Tags: create pie chart Aspose.Cells C# | add chart to new worksheet Aspose.Cells | set chart data source range Aspose.Cells | configure chart size pixels Aspose.Cells | assign chart title Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills cells A1:A5 on a "Data" sheet, adds a "Chart" sheet, inserts a 400x300 pixel pie chart linked to Data!A1:A5, sets the chart title to "Sample Pie Chart", and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample data in A1:A5
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";
            dataSheet.Cells["A1"].PutValue(10);
            dataSheet.Cells["A2"].PutValue(20);
            dataSheet.Cells["A3"].PutValue(30);
            dataSheet.Cells["A4"].PutValue(25);
            dataSheet.Cells["A5"].PutValue(15);

            // Add a new worksheet that will contain the chart
            Worksheet chartSheet = workbook.Worksheets.Add("Chart");

            // Add a pie chart to the new worksheet (position and size in pixels)
            int chartIndex = chartSheet.Charts.Add(ChartType.Pie, 0, 0, 400, 300);
            Chart pieChart = chartSheet.Charts[chartIndex];

            // Set the data source for the chart (range from the Data sheet)
            pieChart.NSeries.Add("Data!A1:A5", true);

            // Optional: set a title for the chart
            pieChart.Title.Text = "Sample Pie Chart";

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
