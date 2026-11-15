// Title: How to enable data labels and change their font size and color on a column chart with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set data label font size on column chart | change data label color in Aspose.Cells chart using .NET | enable data labels and customize appearance in Aspose.Cells column chart | Aspose.Cells example for formatting chart data labels in C# | C# Aspose.Cells chart data labels show values and style
// Tags: Aspose.Cells column chart data label formatting | set data label font size Aspose.Cells .NET | customize data label color Aspose.Cells | activate chart data labels Aspose.Cells | Aspose.Cells chart data label appearance

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds a column chart, enables data labels for the series, sets the label font size to 14 and color to dark blue, and saves the file as ChartWithResizedDataLabels.xlsx.
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

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels (show values)
            series.DataLabels.ShowValue = true;

            // Adjust font size and color for the data labels
            series.DataLabels.Font.Size = 14;
            series.DataLabels.Font.Color = Color.DarkBlue;

            // Save the workbook
            string outputPath = "ChartWithResizedDataLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
