// Title: C# – Format Second Series Data Labels (Bold Red Font, Yellow Background) with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a column chart with two series, enable data labels for the second series, and apply bold red font and a yellow label background using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart formatting | second series data labels | bold red font Aspose.Cells | yellow background chart labels | .NET chart data label style | customize chart series labels | Aspose.Cells example
// Common Searches: Aspose.Cells format data labels second series | C# chart label bold red font Aspose.Cells | set yellow background for chart data labels .NET | how to style specific series labels in Aspose.Cells | apply font and background to chart data labels C#
// Developer Intent: Apply bold red font and a yellow background to the data labels of the second series in a column chart using Aspose.Cells for .NET.
// Use Cases: Highlight forecast values in a sales chart by styling the second series' data labels. | Create presentation‑ready workbooks where specific series stand out visually. | Generate reports that emphasize key metrics with custom label colors and fonts.
// AI Prompts: Generate C# code with Aspose.Cells that sets the second series data labels to bold red font and a yellow background. | Show an example of formatting only one series' data labels in a chart, including font style, color, and label area fill. | Explain how to apply font settings to all child nodes of chart data labels in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart with two series, enable data labels for the second series, and apply bold red font and a yellow label background using Aspose.Cells for .NET.
    public class FormatSecondSeriesDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for two series
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add the two series to the chart
                chart.NSeries.Add("B2:B4", true); // First series
                chart.NSeries.Add("C2:C4", true); // Second series
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the second series
                Series secondSeries = chart.NSeries[1];
                secondSeries.DataLabels.ShowValue = true;

                // Apply bold font and red color to the data labels
                secondSeries.DataLabels.Font.IsBold = true;
                secondSeries.DataLabels.Font.Color = Color.Red;

                // Set yellow background for the data labels
                secondSeries.DataLabels.Area.BackgroundColor = Color.Yellow;

                // Apply the font settings to all child nodes of the data labels
                secondSeries.DataLabels.ApplyFont();

                // Save the workbook
                workbook.Save("FormattedSecondSeriesDataLabels.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            FormatSecondSeriesDataLabels.Run();
        }
    }
}
