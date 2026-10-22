// Title: How to retrieve the worksheet that contains a specific chart using Aspose.Cells Chart.Worksheet property in C#
// AI Prompts: Write C# code that creates a workbook, adds a column chart, and then uses the Chart.Worksheet property to obtain the chart's parent worksheet. | Show how to access the worksheet object of an existing chart in an Aspose.Cells workbook via the Chart.Worksheet member.
// Common Searches: Aspose.Cells C# get parent worksheet of a chart | Chart.Worksheet property example for retrieving worksheet in .NET | How to find which worksheet a chart belongs to using Aspose.Cells | Retrieve worksheet from chart object Aspose.Cells C# code sample
// Tags: Aspose.Cells chart worksheet retrieval | Chart.Worksheet property C# | retrieve parent worksheet from chart Aspose.Cells | Aspose.Cells chart to worksheet mapping

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a workbook, populates data, adds a column chart, and demonstrates how to use the Chart.Worksheet property to get the worksheet that hosts the chart, then outputs the worksheet name and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Name = "SampleChart";

            // Define the data range for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Retrieve the worksheet that contains this chart using Chart.Worksheet
            Worksheet parentWorksheet = chart.Worksheet;

            // Display the name of the worksheet that holds the chart
            Console.WriteLine("Chart is located in worksheet: " + parentWorksheet.Name);

            // Save the workbook to a file
            workbook.Save("ChartDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
