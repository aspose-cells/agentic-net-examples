// Title: Set a Horizontal Two‑Color Linear Gradient Background for a Chart in Aspose.Cells (.NET C#)
// Description: Demonstrates how to use Aspose.Cells for .NET to apply a horizontal two‑color linear gradient (e.g., black‑to‑white) to a chart's plot‑area background via the FillFormat API and save the workbook as an Excel file.
// Keywords: Aspose.Cells chart gradient | linear gradient fill C# | SetTwoColorGradient Aspose | chart plot area background gradient | horizontal gradient Aspose.Cells | Excel chart styling .NET | FillFormat gradient example | gradient variant Aspose.Cells | C# Aspose.Cells chart background | two‑color gradient chart
// Common Searches: how to add a horizontal gradient to a chart background using Aspose.Cells | Aspose.Cells C# set two‑color gradient for chart plot area | linear gradient fill for Excel chart with Aspose.Cells | chart background gradient example Aspose.Cells .NET | SetTwoColorGradient method Aspose.Cells chart
// Developer Intent: Apply a horizontal two‑color linear gradient to a chart’s plot‑area background with Aspose.Cells for .NET.
// Use Cases: Create visually striking reports where chart backgrounds transition between brand colors. | Design presentation‑ready workbooks with pre‑styled charts that use high‑contrast gradients for better readability. | Automate the generation of dashboards that require consistent gradient styling across multiple charts.
// AI Prompts: Generate C# code with Aspose.Cells to apply a vertical three‑color gradient to a chart’s plot area. | Show how to change the gradient variant and direction (e.g., diagonal) for a chart background using Aspose.Cells. | Explain the steps to set a custom color gradient on the chart area (not the plot area) in Aspose.Cells for .NET.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to use Aspose.Cells for .NET to apply a horizontal two‑color linear gradient (e.g., black‑to‑white) to a chart's plot‑area background via the FillFormat API and save the workbook as an Excel file.
class ChartGradientBackgroundExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data for the chart (optional but makes the chart visible)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Item 1");
        sheet.Cells["A3"].PutValue("Item 2");
        sheet.Cells["A4"].PutValue("Item 3");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the fill format of the chart's plot area background
        FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;

        // Set the fill type to gradient to enable gradient properties
        plotAreaFill.FillType = FillType.Gradient;

        // Configure a linear two‑color gradient (horizontal) with contrasting colors
        // Using black and white as an example of high contrast
        plotAreaFill.GradientFill.SetTwoColorGradient(
            Color.Black,               // First color
            Color.White,               // Second color
            GradientStyleType.Horizontal, // Linear gradient direction
            1);                        // Variant (1‑4)

        // Save the workbook with the configured chart background
        workbook.Save("ChartWithLinearGradientBackground.xlsx");
    }
}
