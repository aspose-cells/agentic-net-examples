// Title: Auto‑Resize Chart Data Label Shapes After Applying Bold Font with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart, enable data labels, apply a bold black font, and set IsResizeShapeToFitText to true so label shapes automatically expand and prevent text overflow before saving the file.
// Keywords: Aspose.Cells | C# chart data labels | auto resize data label shape | IsResizeShapeToFitText | .NET Excel chart | prevent label overflow | bold font data labels
// Common Searches: Aspose.Cells auto resize data label after bold | prevent chart data label overflow C# | IsResizeShapeToFitText example | how to fit bold text in Excel chart labels | Aspose.Cells resize label shape
// Developer Intent: Enable automatic resizing of chart data label shapes to accommodate bold text and avoid overflow.
// Use Cases: Generating Excel reports where data label fonts are emphasized without truncating values. | Creating dynamic charts that adapt label size when font styles change at runtime. | Automating workbook creation with column charts that maintain readable labels after styling.
// AI Prompts: Provide C# code using Aspose.Cells to add a column chart, set data label font to bold, and enable IsResizeShapeToFitText. | Explain the purpose of IsResizeShapeToFitText for chart data labels and when it should be applied. | Show a step‑by‑step guide to prevent data label overflow in Excel charts with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelResizeDemo
{
    // Demonstrates how to create a workbook, add a column chart, enable data labels, apply a bold black font, and set IsResizeShapeToFitText to true so label shapes automatically expand and prevent text overflow before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
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
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Enable data labels for the first series
            DataLabels labels = chart.NSeries[0].DataLabels;
            labels.ShowValue = true;

            // Apply bold font to the data labels
            labels.Font.IsBold = true;
            labels.Font.Color = Color.Black;
            labels.Font.Size = 12;

            // After making the font bold, enable auto‑resize of the label shape
            // so the shape expands to contain the larger text and prevents overflow.
            labels.IsResizeShapeToFitText = true;

            // Save the workbook (lifecycle: save)
            workbook.Save("DataLabelResizeAfterBold.xlsx");
        }
    }
}
