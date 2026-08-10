// Title: Resize Chart Data Label Shapes After Text Wrap with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, enables data labels, wraps each point's label text, activates auto‑resize, sets a fixed width, and saves the file so multi‑line labels expand vertically without truncation.
// Keywords: Aspose.Cells | .NET | C# | chart data label | text wrap | auto resize | IsResizeShapeToFitText | IsTextWrapped | column chart | Excel export
// Common Searches: Aspose.Cells enable text wrap for chart data labels | auto resize data label shape after wrap .NET | set data label width Aspose.Cells chart | multi line data labels in Excel using Aspose.Cells | resize chart label shape to fit text
// Developer Intent: Make each chart point’s data label automatically grow in height to accommodate wrapped multi‑line text while keeping a consistent width.
// Use Cases: Generate column charts where long values appear as wrapped, height‑adjusting data labels. | Produce Excel reports with lengthy category names displayed on data labels without clipping. | Create dashboards that maintain uniform label width but variable height based on content length.
// AI Prompts: Show how to enable text wrapping and auto‑resize for individual chart data labels using Aspose.Cells in C#. | Provide C# code that sets a fixed width for data labels and lets the shape expand vertically after wrapping. | Explain the interaction between IsTextWrapped and IsResizeShapeToFitText for chart point labels in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart, enables data labels, wraps each point's label text, activates auto‑resize, sets a fixed width, and saves the file so multi‑line labels expand vertically without truncation.
    public class ResizeDataLabelShapesAfterWrap
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(123);
                sheet.Cells["B3"].PutValue(4567);
                sheet.Cells["B4"].PutValue(8910);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                DataLabels seriesLabels = chart.NSeries[0].DataLabels;
                seriesLabels.ShowValue = true;

                // Apply text wrapping and enable auto‑resize for each point's data label
                foreach (ChartPoint pt in chart.NSeries[0].Points)
                {
                    // Wrap the label text
                    pt.DataLabels.IsTextWrapped = true;

                    // Allow the shape to grow to fit the wrapped text
                    pt.DataLabels.IsResizeShapeToFitText = true;

                    // Set an initial width; the shape will expand vertically as needed
                    pt.DataLabels.Width = 80;
                }

                // Save the workbook
                workbook.Save("ResizeDataLabelsAfterWrap.xlsx");
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
            ResizeDataLabelShapesAfterWrap.Run();
        }
    }
}
