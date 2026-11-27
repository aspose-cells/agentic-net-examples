// Title: How to apply mixed font sizes and colors to a chart point’s data label using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells set different font sizes for parts of a chart data label C# | C# Aspose.Cells customize characters formatting in a chart point label | Create rich text data label for a specific column chart point Aspose.Cells | Change color of month text in chart data label using Aspose.Cells .NET | Use Characters method to style chart point label in Aspose.Cells C# example
// Tags: chart point data label formatting Aspose.Cells | mixed font sizes in chart labels .NET | rich text label characters method C# | customize chart data label colors Aspose.Cells | Aspose.Cells column chart data label styling

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a column chart with sample sales data, enables data labels for the first series, assigns a custom label "Jan120" to the first chart point, and uses the Characters method to apply different font sizes and colors to the month and sales portions before saving the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true; // Show the numeric value

            // Access the first data point in the series
            ChartPoint point = series.Points[0];

            // Set a custom text for the data label (e.g., "Jan120")
            point.DataLabels.Text = "Jan120";

            // Apply mixed font sizes:
            // - Make the month part ("Jan") larger (size 14)
            // - Keep the sales number part ("120") default size (size 10)
            // Characters(startIndex, length) works on zero‑based indices.
            // "Jan120" => indices 0‑2 for "Jan", 3‑5 for "120"
            point.DataLabels.Characters(0, 3).Font.Size = 14; // "Jan"
            point.DataLabels.Characters(3, 3).Font.Size = 10; // "120"

            // Optionally, change colors to highlight the parts
            point.DataLabels.Characters(0, 3).Font.Color = Color.Blue;
            point.DataLabels.Characters(3, 3).Font.Color = Color.DarkGreen;

            // Save the workbook
            string outputPath = "RichTextDataLabel.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
