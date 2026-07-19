// Title: C# – Apply Custom Font to Chart Data Labels and Auto‑Resize Shapes with Aspose.Cells
// Description: Learn how to create a column chart in Aspose.Cells for .NET, set a Calibri 16 pt bold blue font for data labels, propagate the style to all points, and enable automatic shape resizing so labels never get clipped.
// Keywords: Aspose.Cells chart data label font | C# Aspose.Cells custom label style | auto resize data label shape | IsResizeShapeToFitText Aspose | ApplyFont chart series | .NET chart label styling | ChartPoint label auto‑size | Aspose.Cells example
// Common Searches: change font of chart data labels Aspose.Cells C# | auto‑fit data label shape Aspose.Cells chart | set bold blue Calibri font for Aspose.Cells labels | resize data label shapes to fit text Aspose.Cells | Aspose.Cells chart label styling tutorial
// Developer Intent: The developer wants to style chart data labels with a specific font and ensure each label’s shape automatically expands to accommodate the new text.
// Use Cases: Generate a column chart where every data label uses a 16‑pt bold blue Calibri font. | Prevent label clipping by enabling automatic shape resizing after font changes. | Apply font settings to an entire series with a single ApplyFont call.
// AI Prompts: Write C# code using Aspose.Cells to set a custom font for chart data labels and enable shape auto‑fit. | Show how to loop through ChartPoint objects and turn on IsResizeShapeToFitText and IsAutomaticSize for each label. | Explain the effect of ApplyFont() on individual data label shapes in an Aspose.Cells chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsCustomDataLabelFont
{
    // Learn how to create a column chart in Aspose.Cells for .NET, set a Calibri 16 pt bold blue font for data labels, propagate the style to all points, and enable automatic shape resizing so labels never get clipped.
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
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["A5"].PutValue("D");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["B5"].PutValue(40);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Apply a custom font style to the data labels
            series.DataLabels.Font.Name = "Calibri";
            series.DataLabels.Font.Size = 16;
            series.DataLabels.Font.Color = Color.Blue;
            series.DataLabels.Font.IsBold = true;

            // Propagate the font settings to all child label objects
            series.DataLabels.ApplyFont();

            // Ensure each individual data label shape resizes to fit the new font
            foreach (ChartPoint point in series.Points)
            {
                // Enable auto‑fit for the shape that holds the data label
                point.DataLabels.IsResizeShapeToFitText = true;

                // Optionally, you can let the shape auto‑size based on its content
                point.DataLabels.IsAutomaticSize = true;
            }

            // Save the workbook with the customized chart
            workbook.Save("CustomDataLabelFont.xlsx");
        }
    }
}
