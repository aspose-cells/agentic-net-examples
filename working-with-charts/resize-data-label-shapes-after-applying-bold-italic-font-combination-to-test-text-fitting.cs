// Title: Auto‑Resize Chart Data Label Shapes After Bold‑Italic Formatting with Aspose.Cells for .NET
// Description: Demonstrates how to create a column chart in an Excel workbook using Aspose.Cells for .NET, enable data labels, apply bold and italic styling with a dark‑blue color, and automatically resize the label shapes to fit the formatted text by setting IsResizeShapeToFitText and an initial width. The workbook is saved as ResizeDataLabelShapes.xlsx.
// Keywords: Aspose.Cells | C# | .NET | chart data labels | auto resize shape | IsResizeShapeToFitText | bold italic font | column chart | Excel automation | data label width | ResizeDataLabelShapes.xlsx
// Common Searches: Aspose.Cells auto resize chart data label shape | IsResizeShapeToFitText example C# | bold italic data labels Aspose.Cells | set data label width then auto‑fit Aspose.Cells | chart data label shape fitting text Aspose.Cells
// Developer Intent: Enable chart data label shapes to expand automatically so that bold‑italic formatted text fits without being clipped.
// Use Cases: Generate Excel reports with column charts where data labels must remain readable after applying bold‑italic styling. | Create dashboards that programmatically adjust label shapes to accommodate dynamic font changes. | Automate workbook creation where label dimensions are unknown beforehand, relying on IsResizeShapeToFitText to handle sizing.
// AI Prompts: Provide C# code that sets IsResizeShapeToFitText for chart data labels after applying bold and italic fonts using Aspose.Cells. | Show how to define an initial width for data labels and let them auto‑expand to fit the text in Aspise.Cells. | Explain the interaction between Font.IsBold, Font.IsItalic, and IsResizeShapeToFitText for chart data labels.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a column chart in an Excel workbook using Aspose.Cells for .NET, enable data labels, apply bold and italic styling with a dark‑blue color, and automatically resize the label shapes to fit the formatted text by setting IsResizeShapeToFitText and an initial width. The workbook is saved as ResizeDataLabelShapes.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;

        // Apply bold and italic font to the data labels
        dataLabels.Font.IsBold = true;
        dataLabels.Font.IsItalic = true;
        dataLabels.Font.Color = Color.DarkBlue; // optional visual cue

        // Enable auto‑resize of the data label shape to fit the formatted text
        dataLabels.IsResizeShapeToFitText = true;

        // Set an initial small width to demonstrate the auto‑fit behavior
        dataLabels.Width = 40;

        // Save the workbook
        workbook.Save("ResizeDataLabelShapes.xlsx");
    }
}
