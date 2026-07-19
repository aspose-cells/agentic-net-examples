// Title: Aspose.Cells for .NET – Resize Chart Data Label Shapes to Fit Wrapped Multi‑Line Text (C#)
// Description: Demonstrates how to create a column chart with long category names, enable text wrapping on data labels, set a fixed label width, and automatically resize the label shapes so the wrapped content is fully visible. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells C# chart data label wrap | resize data label shape Aspose.Cells | IsTextWrapped property | IsResizeShapeToFitText | WidthPixel chart label | multi‑line data labels .NET | column chart label formatting | Aspose.Cells example GitHub
// Common Searches: Aspose.Cells enable text wrap for chart data labels | auto‑fit data label shape after wrapping | set fixed width for chart data labels C# | resize data label shapes Aspose.Cells .NET | how to make multi‑line data labels in Aspose.Cells
// Developer Intent: Automatically expand chart data label shapes to accommodate wrapped, multi‑line text.
// Use Cases: Generate a column chart where category names are long and need to wrap within a defined width. | Show numeric values on data labels while keeping the label shape sized to the wrapped text. | Create readable multi‑line data labels without manual size calculations by using IsResizeShapeToFitText.
// AI Prompts: Write C# code with Aspose.Cells that adds a column chart, turns on IsTextWrapped for data labels, sets WidthPixel, and enables IsResizeShapeToFitText to auto‑adjust label shapes. | Provide a step‑by‑step explanation of how to make chart data labels wrap and resize automatically in Aspose.Cells for .NET. | Generate a GitHub‑ready Aspose.Cells example that demonstrates multi‑line data labels with automatic shape resizing.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a column chart with long category names, enable text wrapping on data labels, set a fixed label width, and automatically resize the label shapes so the wrapped content is fully visible. The workbook is saved as an XLSX file.
class ResizeDataLabelShapes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with long category names to force wrapping
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Long Category Name 1");
        sheet.Cells["A3"].PutValue("Long Category Name 2");
        sheet.Cells["A4"].PutValue("Long Category Name 3");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Access the data labels of the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;               // Show the numeric values
        dataLabels.Position = LabelPositionType.Center;

        // Enable text wrapping so that long category names become multi‑line
        dataLabels.IsTextWrapped = true;

        // Allow the data label shape to auto‑fit the wrapped text
        dataLabels.IsResizeShapeToFitText = true;

        // Optionally set a fixed width to demonstrate wrapping effect
        dataLabels.WidthPixel = 80;

        // Save the workbook
        workbook.Save("ResizeDataLabelShapes.xlsx");
    }
}
