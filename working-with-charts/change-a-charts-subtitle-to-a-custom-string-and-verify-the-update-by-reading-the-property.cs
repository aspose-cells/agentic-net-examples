// Title: Simulate a chart subtitle in Aspose.Cells by setting Title.Text and read it back in C#
// AI Prompts: Assign a custom string to the chart's Title.Text to act as a subtitle, then output the Title.Text value for confirmation. | Create a column chart, use the Title.Text property as a subtitle placeholder, and programmatically retrieve the text to verify the change. | Set the chart's Title.Text in Aspose.Cells, read the property back, and display it to ensure the subtitle simulation succeeded.
// Common Searches: how to use Title.Text as subtitle for an Aspose.Cells chart in C# | Aspose.Cells chart title verification after setting custom subtitle text | C# example adding subtitle‑like text to an Excel chart with Aspose.Cells
// Tags: Aspose.Cells chart title text property | C# simulate subtitle with chart title Aspose.Cells | column chart custom title text Aspose.Cells | read chart title value C# Aspose.Cells | validate chart title update Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, adds sample data, inserts a column chart, sets the chart's Title.Text to a custom string (used as a subtitle surrogate), reads back the Title.Text to confirm the assignment, and saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data series for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set a custom title for the chart (Aspose.Cells does not have a Subtitle property)
            chart.Title.Text = "Sales Overview Q1";

            // Verify the title by reading the property
            string title = chart.Title.Text;
            Console.WriteLine("Chart title: " + title);

            // Save the workbook to a file (lifecycle rule: save)
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
