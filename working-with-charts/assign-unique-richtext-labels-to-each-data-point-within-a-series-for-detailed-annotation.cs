// Title: C# – Assign Individual Rich‑Text Data Labels to Column‑Chart Points with Aspose.Cells
// Description: This example demonstrates how to create an Excel workbook, add a column chart, bind it to category/value data, enable data labels, turn off automatic text, and then apply a custom multi‑line label to each chart point. Each label shows the category name and its numeric value, uses bold 10‑pt font, and receives a unique color, before the workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells C# chart data labels | custom rich text labels Aspose.Cells | column chart point annotation .NET | disable auto text chart labels | per‑point label color Excel | multi‑line data labels Aspose.Cells | Excel chart customization C# | Aspose.Cells chart formatting
// Common Searches: how to set custom text for each point in an Aspose.Cells chart | Aspose.Cells C# assign different colors to data labels per bar | disable auto‑generated data labels in Aspose.Cells column chart | add multi‑line labels with category and value in Aspose.Cells | rich‑text formatting for chart point labels using Aspose.Cells
// Developer Intent: Create a column chart and give every data point a uniquely formatted rich‑text label.
// Use Cases: Display both category and numeric value in a clear, multi‑line label for each bar. | Emphasize individual columns by applying distinct label colors, improving visual analysis. | Produce Excel reports where default labels are replaced with styled, bold labels that convey extra context.
// AI Prompts: Show how to include the series name in each rich‑text label while preserving custom colors. | Give an example that varies the font size of data labels based on the point's value. | Explain how to export the chart with these custom rich‑text labels to PDF using Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRichTextDataLabels
{
    // This example demonstrates how to create an Excel workbook, add a column chart, bind it to category/value data, enable data labels, turn off automatic text, and then apply a custom multi‑line label to each chart point. Each label shows the category name and its numeric value, uses bold 10‑pt font, and receives a unique color, before the workbook is saved as an .xlsx file.
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

                // Populate sample data (Category / Value)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["A4"].PutValue("Gamma");
                sheet.Cells["B4"].PutValue(45);
                sheet.Cells["A5"].PutValue("Delta");
                sheet.Cells["B5"].PutValue(60);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 22, 15);
                Chart chart = sheet.Charts[chartIdx];

                // Bind data to the chart (values are vertical)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;          // show the numeric value
                series.DataLabels.ShowCategoryName = false; // we will supply custom text

                // Assign a unique rich‑text label to each point
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];

                    // Disable auto‑generated text so we can set our own
                    point.DataLabels.IsAutoText = false;

                    // Retrieve category name from column A
                    string category = sheet.Cells[i + 2, 0].StringValue;

                    // Convert the point's Y value to double
                    double value = Convert.ToDouble(point.YValue);

                    // Set custom label text
                    point.DataLabels.Text = $"Item: {category}\nValue: {value}";

                    // Apply rich‑text formatting: different color per point
                    Color labelColor = i switch
                    {
                        0 => Color.DarkRed,
                        1 => Color.DarkGreen,
                        2 => Color.DarkBlue,
                        _ => Color.Purple
                    };
                    point.DataLabels.Font.Color = labelColor;
                    point.DataLabels.Font.IsBold = true;
                    point.DataLabels.Font.Size = 10;
                }

                // Ensure output directory exists
                string outputPath = "RichTextDataLabels.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
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
