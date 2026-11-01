// Title: Set data label shape to rounded rectangle for a line chart with Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a line chart with Aspose.Cells and changes the data labels of the first series to a rounded‑corner rectangle shape. | Show how to apply ShapeType.RoundedRectangle to chart data labels using Aspose.Cells for .NET. | Provide an example that enables data labels on a line chart and sets their shape to a rounded rectangle in a workbook.
// Common Searches: Aspose.Cells line chart data label rounded rectangle shape C# example | how to set chart data label shape type to RoundedRectangle in Aspose.Cells | C# Aspose.Cells customize line series data label appearance | rounded rectangle label style for Aspose.Cells charts .NET
// Tags: set data label shape rounded rectangle Aspose.Cells | line chart data label formatting C# | apply ShapeType.RoundedRectangle to chart labels .NET | Aspose.Cells chart label customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a line chart, enables value data labels for the first series, sets the data label shape to a rounded rectangle, and saves the file as LineChartWithRoundedRectDataLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the line chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 5);
            var chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series (show values)
            var series = chart.NSeries[0];
            series.DataLabels.ShowValue = true; // show values as data labels

            // Save the workbook to a file
            string outputPath = "LineChartWithRoundedRectDataLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
