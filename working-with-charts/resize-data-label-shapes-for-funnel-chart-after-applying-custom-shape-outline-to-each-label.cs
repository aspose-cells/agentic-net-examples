// Title: How to resize funnel chart data label shapes and add a custom outline using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set fixed width and height for funnel chart data labels | Apply custom border to individual funnel chart data label shapes in .NET | Disable auto‑resize of chart data label shapes for each point with Aspose.Cells | Change background mode of funnel chart data labels to opaque using Aspose.Cells | Resize and style data labels in a funnel chart programmatically with Aspose.Cells
// Tags: funnel chart data label resizing Aspose.Cells | chart data label border styling C# | disable automatic label shape resize Aspose.Cells | explicit label width height .NET | opaque background mode for chart data labels Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace FunnelChartDataLabelResize
{
    // The example creates a workbook, adds a funnel chart with sample data, enables data labels, and for each chart point disables automatic shape resizing, sets a fixed width and height, applies a visible dark‑blue border, and makes the label background opaque before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the funnel chart
                // Column A – Categories, Column B – Values
                sheet.Cells["A1"].PutValue("Stage");
                sheet.Cells["B1"].PutValue("Count");
                sheet.Cells["A2"].PutValue("Prospects");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Qualified Leads");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Proposals");
                sheet.Cells["B4"].PutValue(45);
                sheet.Cells["A5"].PutValue("Negotiations");
                sheet.Cells["B5"].PutValue(20);
                sheet.Cells["A6"].PutValue("Closed Deals");
                sheet.Cells["B6"].PutValue(10);

                // Add a Funnel chart (rows 8‑20, columns 0‑6)
                int chartIndex = sheet.Charts.Add(ChartType.Funnel, 8, 0, 20, 6);
                Chart funnelChart = sheet.Charts[chartIndex];

                // Set the series data (values) and category data (labels)
                funnelChart.NSeries.Add("B2:B6", true);
                funnelChart.NSeries.CategoryData = "A2:A6";

                // Enable data labels and set a basic position
                DataLabels seriesLabels = funnelChart.NSeries[0].DataLabels;
                seriesLabels.ShowValue = true;
                seriesLabels.Position = LabelPositionType.Center;

                // Apply custom outline and resize each individual data label shape
                foreach (ChartPoint point in funnelChart.NSeries[0].Points)
                {
                    // Access the DataLabels object for the specific point
                    DataLabels pointLabel = point.DataLabels;

                    // Disable automatic resizing to fit the text
                    pointLabel.IsResizeShapeToFitText = false;

                    // Set explicit size for the label shape (in points)
                    pointLabel.Width = 80;   // width in points
                    pointLabel.Height = 30;  // height in points

                    // Apply a custom outline (border) to the label shape
                    pointLabel.Border.IsVisible = true;
                    pointLabel.Border.Color = Color.DarkBlue;
                    pointLabel.Border.Weight = WeightType.MediumLine; // line thickness

                    // Optional: make the background opaque to highlight the outline
                    pointLabel.BackgroundMode = BackgroundMode.Opaque;
                }

                // Save the workbook to a file
                string outputPath = "FunnelChartDataLabelsResized.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("Each data label now has a custom outline and a fixed size.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
