// Title: Assign custom colors to each slice of a pie chart using Aspose.Cells for .NET
// AI Prompts: Create a workbook, add a pie chart, enable IsColorVaried, and set a distinct ForegroundColor for each point from a Color array in C#. | Generate an Excel file with a pie chart where slice colors are mapped to importance levels by assigning colors to chart.NSeries[0].Points. | Show how to programmatically customize individual pie slice colors in Aspose.Cells without relying on the default palette.
// Common Searches: Aspose.Cells set individual slice colors in a pie chart C# | C# how to use IsColorVaried for pie chart series in Aspose.Cells | assign custom foreground colors to chart points Aspose.Cells .NET example | change pie chart slice colors based on data importance using Aspose.Cells | create Excel pie chart with specific colors for each category Aspose.Cells
// Tags: pie chart custom slice colors Aspose.Cells | IsColorVaried property C# chart | assign point foreground color Aspose.Cells | Excel pie chart color mapping .NET | generate colored pie chart workbook Aspose | color variation per data point Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills category data, adds a pie chart, enables varied colors for the series, and applies a predefined Red‑Orange‑Green‑Blue color array to each slice before saving the file as PieChartCustomColors.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data: categories and their corresponding values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            string[] categories = { "A", "B", "C", "D" };
            double[] values = { 30, 20, 25, 25 };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue(values[i]);    // Column B
            }

            // Add a pie chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 25, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B5", true);

            // Allow each slice to have its own color
            chart.NSeries[0].IsColorVaried = true;

            // Define custom colors for each slice based on importance
            Color[] sliceColors = { Color.Red, Color.Orange, Color.Green, Color.Blue };

            // Apply the custom colors to the individual points (slices)
            for (int i = 0; i < sliceColors.Length; i++)
            {
                chart.NSeries[0].Points[i].Area.ForegroundColor = sliceColors[i];
            }

            // Save the workbook with the customized pie chart
            workbook.Save("PieChartCustomColors.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
