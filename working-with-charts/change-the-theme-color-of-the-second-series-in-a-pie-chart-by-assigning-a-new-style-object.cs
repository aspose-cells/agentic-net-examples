// Title: Set a custom theme color for the second series of a pie chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change color of second series in a pie chart | apply solid fill style to specific chart series with Aspose.Cells .NET | how to modify chart series style programmatically in Excel using Aspose.Cells | set red fill for second series of a pie chart in C# Aspose.Cells example | assign new Style object to chart series in Aspose.Cells workbook
// Tags: modify chart series style Aspose.Cells .NET | set series fill color pie chart C# | apply custom Style object to Excel chart series | change second series theme color Aspose.Cells | programmatic chart series styling Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The code loads or creates a workbook, ensures a pie chart with at least two series exists, creates a Style with a solid red fill, assigns this Style to the second series, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                // Add a default worksheet and a sample pie chart for demonstration.
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";

                // Sample data for the chart.
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(30);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(70);

                // Create a pie chart.
                int chartIndex = ws.Charts.Add(ChartType.Pie, 5, 0, 25, 10);
                Chart chart = ws.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries[0].CategoryData = "A2:A3";
            }

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart.
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart (assumed to be a pie chart) and enforce its type.
            Chart chartObj = worksheet.Charts[0];
            chartObj.Type = ChartType.Pie;

            // Ensure the chart has at least two series.
            if (chartObj.NSeries.Count < 2)
            {
                Console.WriteLine("The chart does not contain a second series to style.");
                return;
            }

            // Get the second series (index 1) of the chart.
            ChartSeries secondSeries = chartObj.NSeries[1];

            // Create a new style with a solid red fill.
            Style newStyle = workbook.CreateStyle();
            newStyle.ForegroundColor = Color.Red;
            newStyle.Pattern = BackgroundType.Solid;

            // Assign the new style to the second series.
            secondSeries.Style = newStyle;

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
