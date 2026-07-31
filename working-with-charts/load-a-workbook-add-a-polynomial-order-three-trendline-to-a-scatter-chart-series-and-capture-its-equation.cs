// Title: Add a 3rd‑order polynomial trendline to a scatter chart and display its equation using Aspose.Cells for .NET (C#)
// Description: Load or create a workbook, insert X/Y data, build a scatter chart, add a polynomial trendline of order 3, show its equation and R‑squared, customize name and color, save the file, reload it, and verify the trendline settings with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | scatter chart | polynomial trendline | order 3 | trendline equation | R squared | chart customization | Excel workbook | trendline properties
// Common Searches: Aspose.Cells add polynomial trendline C# | display trendline equation Aspose.Cells | set trendline order 3 Aspose.Cells | customize trendline color Aspose.Cells | read trendline properties after saving Aspose.Cells | scatter chart with polynomial fit Aspose.Cells
// Developer Intent: Create a scatter chart, attach a third‑order polynomial trendline, show its equation and R‑squared on the chart, customize appearance, and confirm the settings after saving the workbook.
// Use Cases: Generate Excel reports with fitted polynomial curves for scientific data analysis. | Show trendline equation and R‑squared directly on a chart for presentations or dashboards. | Programmatically set a custom name and red color for a trendline to match corporate branding. | Automated testing to ensure trendline order and display flags persist across workbook saves.
// AI Prompts: Generate C# code that uses Aspose.Cells to add a 3rd‑order polynomial trendline to a scatter chart and enable equation and R‑squared display. | Show how to reload a saved workbook and retrieve the trendline order, DisplayEquation, and DisplayRSquared flags with Aspose.Cells. | Explain how to assign a custom name and red color to a polynomial trendline in an Aspose.Cells chart. | Provide a step‑by‑step guide for creating sample X/Y data, building a scatter chart, and applying a polynomial trendline in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Load or create a workbook, insert X/Y data, build a scatter chart, add a polynomial trendline of order 3, show its equation and R‑squared, customize name and color, save the file, reload it, and verify the trendline settings with Aspose.Cells for .NET.
class PolynomialTrendlineExample
{
    static void Main()
    {
        // Load an existing workbook or create a new one if the file does not exist
        Workbook workbook;
        const string inputPath = "input.xlsx";
        if (System.IO.File.Exists(inputPath))
        {
            workbook = new Workbook(inputPath);
        }
        else
        {
            workbook = new Workbook();
        }

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Prepare sample data for a scatter chart
        // -------------------------------------------------
        // X values
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["A5"].PutValue(4);
        // Y values
        sheet.Cells["B1"].PutValue("Y");
        sheet.Cells["B2"].PutValue(2);
        sheet.Cells["B3"].PutValue(5);
        sheet.Cells["B4"].PutValue(9);
        sheet.Cells["B5"].PutValue(16);

        // -------------------------------------------------
        // Add a scatter chart
        // -------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the series data (Y values) and X values
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries[0].XValues = "A2:A5";
        chart.NSeries[0].Name = "Sample Series";

        // -------------------------------------------------
        // Add a polynomial trendline of order 3
        // -------------------------------------------------
        int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Polynomial);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

        // Set the polynomial order to 3
        trendline.Order = 3;

        // Show the equation and R‑squared value on the chart
        trendline.DisplayEquation = true;
        trendline.DisplayRSquared = true;

        // Optional: give the trendline a name and color
        trendline.Name = "3rd Order Polynomial";
        trendline.Color = Color.Red;

        // -------------------------------------------------
        // Save the workbook with the new chart and trendline
        // -------------------------------------------------
        const string outputPath = "output.xlsx";
        workbook.Save(outputPath);

        // -------------------------------------------------
        // Reload the workbook and capture the trendline equation
        // -------------------------------------------------
        Workbook loadedWb = new Workbook(outputPath);
        Chart loadedChart = loadedWb.Worksheets[0].Charts[0];
        Trendline loadedTrendline = loadedChart.NSeries[0].TrendLines[0];

        // Aspose.Cells does not expose the equation string directly.
        // The equation is displayed on the chart because DisplayEquation = true.
        // Here we simply confirm that the property is set.
        Console.WriteLine("Trendline Order: " + loadedTrendline.Order);
        Console.WriteLine("Display Equation: " + loadedTrendline.DisplayEquation);
        Console.WriteLine("Display R‑Squared: " + loadedTrendline.DisplayRSquared);
        Console.WriteLine("Trendline added and equation displayed on the chart.");
    }
}
