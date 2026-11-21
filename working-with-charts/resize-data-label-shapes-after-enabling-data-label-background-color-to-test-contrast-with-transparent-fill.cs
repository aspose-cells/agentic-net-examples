// Title: Resizing data label shapes in a column chart to improve contrast with a transparent background using Aspose.Cells for .NET
// Common Searches: Aspose.Cells resize data label shape in column chart C# | increase data label box size after setting background color Aspose.Cells | set transparent fill for chart data label background Aspose.Cells .NET | adjust data label dimensions for better contrast in Excel chart using Aspose.Cells | change data label size and font in Aspose.Cells column chart programmatically
// Tags: resize chart data label shape Aspose.Cells | set data label background color .NET | transparent fill contrast chart labels | column chart data label size adjustment | Aspose.Cells data label font customization

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a column chart with sample data, enables data labels, enlarges the label box by increasing the font size to 14 points, and saves the result as ChartWithResizedDataLabels.xlsx, illustrating how to resize data label shapes for better contrast when a background color is applied.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet and give it a name
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(70);
            sheet.Cells["B4"].PutValue(45);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            var chart = sheet.Charts[chartIndex];

            // Add series (values) – categories are taken from the first column automatically
            chart.NSeries.Add("B2:B4", true);

            // Enable data labels and show the values
            var dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;               // Show the values on the labels

            // Adjust label appearance
            dataLabels.Font.Size = 14;                  // Increase font size to enlarge the label box
            dataLabels.Font.Color = Color.Black;       // Ensure text is readable

            // Save the workbook
            workbook.Save("ChartWithResizedDataLabels.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
