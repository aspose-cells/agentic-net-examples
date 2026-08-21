// Title: Add a 3rd‑Order Polynomial Trendline to a Scatter Chart and Retrieve Its Equation with Aspose.Cells (C#)
// Description: Load a workbook, create a scatter chart, add a third‑order polynomial trendline to the first series, enable equation display, capture the equation from the trendline's DataLabels, output it, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# scatter chart | polynomial trendline order 3 | retrieve trendline equation | chart DataLabels Aspose.Cells | Excel regression equation C# | add trendline Aspose.Cells | trendline equation extraction
// Common Searches: Aspose.Cells add polynomial trendline C# | how to get trendline equation from Aspose.Cells chart | display polynomial regression equation in Excel using Aspose | C# scatter chart trendline order 3 Aspose.Cells | extract trendline equation programmatically
// Developer Intent: Add a third‑order polynomial trendline to a scatter chart series and programmatically obtain its equation using Aspose.Cells for .NET.
// Use Cases: Generate a regression formula directly from workbook data for statistical analysis. | Show the polynomial equation on a chart while also using the formula in further calculations. | Log or export the captured equation string for reporting or downstream processing.
// AI Prompts: Provide C# code with Aspose.Cells that adds a 3rd‑order polynomial trendline to a scatter chart and returns the equation as a string. | Show how to enable equation display on a trendline and read the equation from DataLabels in Aspose.Cells. | Give a step‑by‑step example of creating a scatter chart, adding data ranges, applying a polynomial trendline of order 3, capturing the equation, and saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Load a workbook, create a scatter chart, add a third‑order polynomial trendline to the first series, enable equation display, capture the equation from the trendline's DataLabels, output it, and save the workbook using Aspose.Cells for .NET.
class AddPolynomialTrendline
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or create one if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Add a scatter chart (if the sheet does not already contain one)
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Define data for the scatter chart (example data)
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);
        sheet.Cells["A3"].PutValue(3);
        sheet.Cells["A4"].PutValue(4);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["B2"].PutValue(4);
        sheet.Cells["B3"].PutValue(6);
        sheet.Cells["B4"].PutValue(8);

        // Set the series data (X values and Y values)
        chart.NSeries.Add("B1:B4", true);          // Y values
        chart.NSeries[0].XValues = "A1:A4";       // X values
        chart.NSeries[0].Name = "Sample Series";

        // Add a polynomial trendline (order 3) to the first series
        int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Polynomial);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

        // Set the polynomial order to 3
        trendline.Order = 3;

        // Show the equation (and optionally R‑squared) on the chart
        trendline.DisplayEquation = true;
        trendline.DisplayRSquared = false;

        // Capture the equation text (the equation is stored in the DataLabels of the trendline)
        // Note: Aspose.Cells stores the displayed text in the DataLabels.Text property.
        string equation = trendline.DataLabels.Text;

        // Output the captured equation to the console
        Console.WriteLine("Polynomial (order 3) Trendline Equation: " + equation);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
