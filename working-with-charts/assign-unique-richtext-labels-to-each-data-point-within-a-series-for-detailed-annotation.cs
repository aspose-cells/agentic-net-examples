// Title: C# Aspose.Cells: Add Custom Rich‑Text Data Labels to Each Column Chart Point
// Description: Shows how to generate an Excel workbook, insert a column chart, bind category/value ranges, enable data labels, turn off automatic text, and assign a unique rich‑text label to every chart point. Each label displays the category name and its numeric value with point‑specific font color, boldness, and size, then saves the file as RichTextDataLabels.xlsx.
// Keywords: Aspose.Cells | C# | .NET | column chart | custom data labels | rich text chart labels | chart point formatting | disable auto text | Excel chart labeling | programmatic chart customization
// Common Searches: Aspose.Cells set individual data label text per point | C# add rich‑text labels to chart series Aspose.Cells | change font color of single chart data label .NET | disable auto generated data labels Aspose.Cells | customize Excel chart point labels programmatically
// Developer Intent: Programmatically assign distinct rich‑text labels with custom formatting to each point of a chart series using Aspose.Cells for .NET.
// Use Cases: Display both category and value on each column for clearer reporting. | Highlight specific columns with bold blue labels while using green for others to draw visual attention. | Create detailed annotations per data point for presentation‑ready Excel files. | Generate dynamic dashboards where label styling varies based on business rules.
// AI Prompts: Provide C# code to set a unique rich‑text data label for every point in an Aspose.Cells column chart. | How can I format individual chart point labels with different colors, boldness, and font sizes using Aspose.Cells for .NET? | Explain the steps to disable auto‑generated data label text and assign custom text to each series point in a chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRichTextLabels
{
    // Shows how to generate an Excel workbook, insert a column chart, bind category/value ranges, enable data labels, turn off automatic text, and assign a unique rich‑text label to every chart point. Each label displays the category name and its numeric value with point‑specific font color, boldness, and size, then saves the file as RichTextDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (Category and Value)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(45);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Bind data to the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;        // Show the numeric value
                series.DataLabels.IsAutoText = false;     // Allow custom text per point

                // Assign a unique rich‑text label to each data point
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];

                    // Disable auto‑generated text for this point
                    point.DataLabels.IsAutoText = false;

                    // Retrieve category name
                    string category = sheet.Cells[$"A{i + 2}"].StringValue;

                    // Retrieve numeric value (cast from object)
                    double value = Convert.ToDouble(point.YValue);

                    // Set custom label text
                    point.DataLabels.Text = $"[{category}] = {value:F1}";

                    // Apply individual formatting (rich‑text effect)
                    point.DataLabels.Font.Color = (i % 2 == 0) ? Color.Blue : Color.DarkGreen;
                    point.DataLabels.Font.IsBold = (i % 2 == 0);
                    point.DataLabels.Font.Size = 10 + i; // Vary size slightly
                }

                // Save the workbook
                string outputPath = "RichTextDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
