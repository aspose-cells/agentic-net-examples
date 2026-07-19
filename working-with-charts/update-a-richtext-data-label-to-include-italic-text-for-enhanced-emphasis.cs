// Title: C# – Apply Italic Font to Chart Data Labels (Rich‑Text) Using Aspose.Cells
// Description: This example creates a workbook, inserts a column chart, activates data labels for the first series, and formats those labels with an italic typeface. The ApplyFont method propagates the style to all rich‑text nodes before saving the file as an Excel workbook.
// Keywords: Aspose.Cells chart label italic | C# rich text data label formatting | Excel chart label styling Aspose | set italic font for series labels | ApplyFont rich‑text Aspose.Cells
// Common Searches: how to make chart labels italic in Aspose.Cells C# | rich text formatting for Excel chart data labels | Aspose.Cells apply font style to data labels | C# code to italicize chart series labels | Aspose.Cells ApplyFont example
// Developer Intent: Render chart data labels in an italic typeface programmatically.
// Use Cases: Generate sales dashboards where label values need visual emphasis. | Produce branded reports that require consistent italic styling across all chart labels. | Automate Excel workbook creation with pre‑styled chart annotations for downstream consumers.
// AI Prompts: Show C# code that sets an italic style on Aspose.Cells chart data labels. | Explain how ApplyFont updates child rich‑text nodes of a chart’s data labels. | Provide a snippet to italicize data labels for multiple series in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRichTextDataLabel
{
    // This example creates a workbook, inserts a column chart, activates data labels for the first series, and formats those labels with an italic typeface. The ApplyFont method propagates the style to all rich‑text nodes before saving the file as an Excel workbook.
    public class UpdateDataLabelItalic
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Set the data label font to italic
                series.DataLabels.Font.IsItalic = true;

                // Apply the font settings to all child data label nodes (rich‑text)
                series.DataLabels.ApplyFont();

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RichTextDataLabelItalic.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            UpdateDataLabelItalic.Run();
        }
    }
}
