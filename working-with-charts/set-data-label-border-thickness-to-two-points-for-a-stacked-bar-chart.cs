// Title: Create a stacked horizontal bar chart and apply a 2‑point black border to its first series data labels with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a stacked horizontal bar chart in a new workbook, adds two series, and sets the first series' data label border weight to 2 points with a black color using Aspose.Cells. | Write a C# example that enables data labels on the first series of a stacked bar chart and customizes the label border to be visible, 2 pt thick, and black, then saves the workbook as .xlsx. | Provide a complete Aspose.Cells script that creates sample data, inserts a stacked bar chart, and configures the data label border thickness and color for the first series.
// Common Searches: aspnet set data label border thickness for stacked bar chart Aspose.Cells | C# Aspose.Cells how to change data label border weight in bar chart | example of customizing data label borders in a stacked horizontal bar chart using Aspose.Cells | Aspose.Cells chart data label formatting border size 2 points | set data label border color and thickness in Excel chart with Aspose.Cells .NET
// Tags: Aspose.Cells stacked bar chart data label border | C# set data label border weight Aspose.Cells | Excel chart label border thickness .NET | Aspose.Cells data label formatting example | horizontal stacked bar chart Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds sample data, inserts a horizontal stacked bar chart with two series, enables data labels for the first series, and customizes the label border to be visible, black, and 2 pt thick before saving the file as an .xlsx workbook.
class StackedBarDataLabelBorderDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked bar chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(50);
            sheet.Cells["B4"].PutValue(20);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["C3"].PutValue(40);
            sheet.Cells["C4"].PutValue(30);

            // Add a stacked bar chart (horizontal stacked bar)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels for the first series
            Series series = chart.NSeries[0];
            DataLabels labels = series.DataLabels;
            labels.ShowValue = true;               // Show the value on each data label
            labels.Border.IsVisible = true;        // Ensure the border is drawn
            labels.Border.WeightPt = 2.0;           // Set border thickness to two points
            labels.Border.Color = Color.Black;     // Optional: set border color

            // Determine output file path
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "StackedBarDataLabelBorderDemo.xlsx");

            // Save the workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
