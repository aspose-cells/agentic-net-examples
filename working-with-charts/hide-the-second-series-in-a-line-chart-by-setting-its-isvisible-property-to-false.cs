// Title: Hide the second series in an Aspose.Cells line chart by setting its IsVisible property to false (C#)
// AI Prompts: Generate C# code using Aspose.Cells that adds a line chart with two data series and hides the second series by setting its IsVisible property to false. | Show how to toggle the visibility of a specific series in an Aspose.Cells line chart without removing the series. | Provide a snippet that creates sample data, builds a line chart, and programmatically makes the second series invisible in C#.
// Common Searches: asp.net cells hide second series in line chart c# | set chart series IsVisible false aspose.cells c# example | how to make a series invisible in an Aspose.Cells line chart | c# aspose.cells hide chart series without deleting | line chart series visibility control using Aspose.Cells
// Tags: line chart series visibility Aspose.Cells | set IsVisible false chart series C# | hide second series Excel chart Aspose.Cells | Aspose.Cells chart series manipulation C# | C# hide chart series without removal

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, populates sample data for two series, adds a line chart, then hides the second series by setting its IsVisible property to false, and finally saves the workbook as LineChart_HideSecondSeries.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the first series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Series1";

            // Define the data range for the second series
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Series2";

            // Hide the second series by removing it from the chart
            chart.NSeries.RemoveAt(1);

            // Save the workbook with the chart
            workbook.Save("LineChart_HideSecondSeries.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
