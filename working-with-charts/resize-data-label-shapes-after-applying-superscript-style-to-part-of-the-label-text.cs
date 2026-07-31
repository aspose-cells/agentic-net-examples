// Title: Resize Chart Data Labels After Adding Superscript in Aspose.Cells for .NET
// Description: This example shows how to create a column chart, add data labels, apply a superscript character to part of each label, and force the label shape to resize by toggling the IsResizeShapeToFitText property. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells chart data label resize | superscript formatting chart label C# | IsResizeShapeToFitText toggle | ChartPoint DataLabels font style | adjust label shape after text change
// Common Searches: Aspose.Cells resize data label after superscript | C# chart label superscript shape fit | force chart data label to auto‑size Aspose.Cells | how to apply superscript to chart data label .NET | chart data label size issue after font change
// Developer Intent: Automatically adjust the size of a chart data label after applying superscript formatting to part of its text.
// Use Cases: Display numeric values with a superscript sign (e.g., 10⁺) while keeping the label box correctly sized. | Create mixed‑format data labels (regular and superscript) that adapt to content length. | Generate Excel reports where special characters require font style changes without manual label resizing.
// AI Prompts: Give C# code that resizes a chart data label after setting a superscript character using Aspose.Cells. | Explain why toggling IsResizeShapeToFitText from false to true updates the label size after font changes. | Show how to apply superscript to the last character of a data label and then recalculate the label shape dimensions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example shows how to create a column chart, add data labels, apply a superscript character to part of each label, and force the label shape to resize by toggling the IsResizeShapeToFitText property. The workbook is then saved as an Excel file.
    public class ResizeDataLabelAfterSuperscript
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
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

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Iterate through each point to customize its data label
                foreach (ChartPoint point in series.Points)
                {
                    // Example label: "10⁺" where the plus sign is superscript
                    point.DataLabels.Text = $"{point.YValue}+";

                    // Apply superscript style to the last character ('+')
                    int superscriptStart = point.DataLabels.Text.Length - 1;
                    point.DataLabels.Characters(superscriptStart, 1).Font.IsSuperscript = true;

                    // Force shape resize to fit modified text
                    point.DataLabels.IsResizeShapeToFitText = false;
                    point.DataLabels.IsResizeShapeToFitText = true;
                }

                // Save the workbook
                string outputPath = "ResizeDataLabelAfterSuperscript.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelAfterSuperscript.Run();
        }
    }
}
