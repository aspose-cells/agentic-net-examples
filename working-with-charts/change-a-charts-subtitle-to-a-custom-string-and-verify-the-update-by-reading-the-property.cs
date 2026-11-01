// Title: Set a custom subtitle for a column chart and read it back using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a column chart, assign a custom subtitle string to the chart title, then output the subtitle value to the console. | Modify the subtitle of an existing column chart in an Excel file with Aspose.Cells, retrieve the Title.Text property, and display it. | Generate sample data, bind it to a column chart, set a custom subtitle, verify the change by reading the chart's title, and save the workbook.
// Common Searches: Aspose.Cells C# change chart subtitle and get title text | how to read chart title after setting subtitle with Aspose.Cells .NET | example of setting column chart subtitle programmatically in Excel using Aspose.Cells | retrieve chart Title.Text property after modification in C# | set and verify chart subtitle in Aspose.Cells workbook
// Tags: set chart subtitle Aspose.Cells | read chart title text .NET | column chart subtitle manipulation | Aspose.Cells chart title verification | Excel workbook chart subtitle example

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Demonstrates creating a workbook, adding a column chart, setting a custom subtitle via the chart title, reading the Title.Text property to verify the change, and saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart title (used here as a subtitle substitute)
            chart.Title.IsVisible = true;
            chart.Title.Text = "Custom subtitle for the chart";

            // Verify the update by reading the title property
            string currentTitle = chart.Title.Text;
            Console.WriteLine("Chart title (used as subtitle): " + currentTitle);

            // Save the workbook
            string outputPath = "ChartWithSubtitle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
