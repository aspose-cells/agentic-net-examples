// Title: Create a mixed‑font rich‑text data label for a specific chart point using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that adds a column chart, enables a data label on the first point, and formats the label text so that the word “Value:” appears in 12‑pt blue font while the numeric value appears in 20‑pt red bold font. | Show how to use the Characters method to apply different font sizes, colors, and bold styling to separate character ranges within a chart point’s data label in Aspose.Cells. | Provide the complete example that saves the workbook after applying the rich‑text formatting to the chart point label.
// Common Searches: Aspose.Cells C# format part of a chart data label with different font sizes | how to apply rich text styling to a single point label in an Aspose.Cells column chart | set blue text for label prefix and red bold number in Aspose.Cells chart label | C# Aspose.Cells mixed font colors and sizes in chart point data label example
// Tags: apply mixed font sizes to chart point label Aspose.Cells | character range formatting in Aspose.Cells chart data label | set font color and bold for chart label segment Aspose.Cells | column chart data label rich text styling Aspose.Cells | custom label text for individual chart point C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRichTextDataLabel
{
    // The example creates a workbook, adds a column chart with sample data, enables a data label on the first chart point, and uses the Characters method to apply different font sizes, colors, and bold styling to specific text segments before saving the file as RichTextDataLabel.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Access the first series and its first point
            Series series = chart.NSeries[0];
            ChartPoint point = series.Points[0];

            // Enable data label for this point and set custom text
            DataLabels dataLabel = point.DataLabels;
            dataLabel.ShowValue = true;
            dataLabel.Text = "Value:10";

            // Apply mixed font sizes:
            // "Value:" (characters 0-5) – size 12, blue
            dataLabel.Characters(0, 5).Font.Size = 12;
            dataLabel.Characters(0, 5).Font.Color = Color.Blue;

            // "10" (characters 6-7) – size 20, red, bold
            dataLabel.Characters(6, 2).Font.Size = 20;
            dataLabel.Characters(6, 2).Font.Color = Color.Red;
            dataLabel.Characters(6, 2).Font.IsBold = true;

            // Apply the font settings to all child nodes of the data label
            dataLabel.ApplyFont();

            // Save the workbook
            workbook.Save("RichTextDataLabel.xlsx");
        }
    }
}
