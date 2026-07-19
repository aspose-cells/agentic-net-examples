// Title: Bold and Center‑Aligned Chart Data Labels with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, enables data labels, then formats those labels with bold, black, 12‑pt font and centers the text horizontally before saving the file.
// Keywords: Aspose.Cells .NET chart data labels | bold font chart labels C# | center align Excel chart labels | format chart data labels Aspose | apply font style to Excel chart | C# Aspose.Cells example | Excel column chart styling | text alignment chart labels | font weight chart data labels | Aspose.Cells chart customization
// Common Searches: How to make chart data labels bold using Aspose.Cells | Center align data labels on a column chart in C# | Set font size and color for Excel chart labels with Aspose.Cells | Aspose.Cells format chart data labels | C# example for styling chart labels in Excel
// Developer Intent: The developer wants to style chart data labels—applying bold weight, specific color/size, and horizontal centering—via Aspose.Cells for .NET.
// Use Cases: Produce a sales dashboard where column‑chart values stand out with bold, centered labels. | Generate financial reports that require uniformly styled chart labels for presentation quality. | Automate creation of Excel workbooks with multiple charts that share the same label formatting rules.
// AI Prompts: Give C# code that sets chart data labels to bold, black, 12‑pt font and centers the text using Aspose.Cells. | Explain how to apply font styling and horizontal alignment to all data label nodes of a chart series in Aspose.Cells for .NET. | Show an Aspose.Cells example that customizes data label appearance (weight, color, size, alignment) for a column chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart, enables data labels, then formats those labels with bold, black, 12‑pt font and centers the text horizontally before saving the file.
    public class DataLabelsBoldCenteredDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Apply bold font style to data labels
                series.DataLabels.Font.IsBold = true;
                // Optionally change the font color or size
                series.DataLabels.Font.Color = Color.Black;
                series.DataLabels.Font.Size = 12;

                // Center align the text of data labels horizontally
                series.DataLabels.TextHorizontalAlignment = TextAlignmentType.Center;

                // Apply the font settings to all child label nodes
                series.DataLabels.ApplyFont();

                // Save the workbook to a file
                workbook.Save("DataLabelsBoldCenteredDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DataLabelsBoldCenteredDemo.Run();
        }
    }
}
