// Title: Aspose.Cells C# – Apply Custom Font and Auto‑Resize Data Labels in a Chart
// Description: Demonstrates how to create a workbook, add a column chart, enable data labels, set a bold blue Calibri font (size 16) for the labels, and automatically resize the label shapes to fit the new text using the IsResizeShapeToFitText property. The same styling is applied to each chart point and propagated to child label nodes before saving the file.
// Keywords: Aspose.Cells | C# chart data labels | custom font data labels | auto resize data label shape | IsResizeShapeToFitText | ApplyFont method | column chart Aspose.Cells | .NET Excel chart styling
// Common Searches: Aspose.Cells change data label font C# | auto fit data label shape Aspose.Cells chart | set bold blue Calibri font for chart labels .NET | IsResizeShapeToFitText example Aspose.Cells | apply font to all data label points Aspose.Cells
// Developer Intent: Set a custom font for chart data labels and enable automatic shape resizing so the labels fit the new style.
// Use Cases: Create a column chart, enable data labels, and apply a blue bold Calibri font (size 16) to the series label. | Iterate through each ChartPoint to assign identical font settings and activate IsResizeShapeToFitText for individual point labels. | Call series.DataLabels.ApplyFont() to propagate font changes to any nested label elements before saving the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that sets a custom font for chart data labels and turns on auto‑resize of label shapes. | Show how to apply the same font settings to every point’s data label in an Aspose.Cells chart and ensure the labels auto‑fit the text. | Explain the purpose of IsResizeShapeToFitText and how ApplyFont works when customizing chart data labels in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelFontDemo
{
    // Demonstrates how to create a workbook, add a column chart, enable data labels, set a bold blue Calibri font (size 16) for the labels, and automatically resize the label shapes to fit the new text using the IsResizeShapeToFitText property. The same styling is applied to each chart point and propagated to child label nodes before saving the file.
    class Program
    {
        static void Main()
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

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Apply custom font style to the data labels
            series.DataLabels.Font.Name = "Calibri";
            series.DataLabels.Font.Size = 16;
            series.DataLabels.Font.Color = Color.Blue;
            series.DataLabels.Font.IsBold = true;

            // Ensure the data label shape resizes to fit the new font
            series.DataLabels.IsResizeShapeToFitText = true;

            // Also apply the same settings to each individual point's data label
            foreach (ChartPoint point in series.Points)
            {
                point.DataLabels.Font.Name = "Calibri";
                point.DataLabels.Font.Size = 16;
                point.DataLabels.Font.Color = Color.Blue;
                point.DataLabels.Font.IsBold = true;
                point.DataLabels.IsResizeShapeToFitText = true;
            }

            // Apply the font settings to all child nodes of the data labels
            series.DataLabels.ApplyFont();

            // Save the workbook
            workbook.Save("DataLabelsCustomFontAndResize.xlsx");
        }
    }
}
