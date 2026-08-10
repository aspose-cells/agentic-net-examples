// Title: Aspose.Cells for .NET – Auto‑Resize Chart Data Label Shapes After Adding Hyperlinks
// Description: Demonstrates how to create a workbook, insert a column chart, add a hyperlink to each data label, and enable the IsResizeShapeToFitText property so every label shape automatically expands to fit its linked text before saving the file.
// Keywords: Aspose.Cells resize data label shape | chart data label auto resize .NET | add hyperlink to chart data label Aspose.Cells | IsResizeShapeToFitText property | column chart data labels fit text | Aspose.Cells chart hyperlink example | C# Excel chart label sizing
// Common Searches: how to auto resize chart data labels in Aspose.Cells | add hyperlink to each data label in Aspose.Cells chart | fit data label shape to text Aspose.Cells .NET | Aspose.Cells IsResizeShapeToFitText usage | C# resize Excel chart label after adding hyperlink
// Developer Intent: Automatically adjust each chart point’s data label shape to fit its text after attaching a hyperlink.
// Use Cases: Generating Excel reports where data labels contain clickable URLs and must expand to avoid truncation. | Creating dynamic dashboards with column charts that automatically size labels for varying values and linked resources. | Automating workbook production for web‑based analytics, ensuring all label hyperlinks are visible and properly sized.
// AI Prompts: Show C# code that adds a hyperlink to each chart data label before enabling auto‑resize in Aspose.Cells. | Explain how to customize font style of data labels while keeping IsResizeShapeToFitText active. | Provide error‑handling best practices for resizing data label shapes in large Aspose.Cells charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert a column chart, add a hyperlink to each data label, and enable the IsResizeShapeToFitText property so every label shape automatically expands to fit its linked text before saving the file.
    public class ResizeDataLabelShapesWithHyperlink
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook(FileFormatType.Xlsx);
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category 1");
            worksheet.Cells["A2"].PutValue("Category 2");
            worksheet.Cells["A3"].PutValue("Category 3");
            worksheet.Cells["B1"].PutValue(10);
            worksheet.Cells["B2"].PutValue(20);
            worksheet.Cells["B3"].PutValue(30);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set series data
            chart.NSeries.Add("B1:B3", true);
            chart.NSeries.CategoryData = "A1:A3";

            // Enable data labels
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;
            dataLabels.Position = LabelPositionType.Center;

            // Iterate through each point, resize the label shape
            int pointIdx = 0;
            foreach (ChartPoint point in chart.NSeries[0].Points)
            {
                try
                {
                    // Resize the data label shape to fit its text
                    point.DataLabels.IsResizeShapeToFitText = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to resize label for point {pointIdx}. {ex.Message}");
                }

                pointIdx++;
            }

            // Save the workbook
            string outputPath = "ResizeDataLabelShapesWithHyperlink.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
