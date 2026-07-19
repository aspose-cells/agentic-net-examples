// Title: Add a Cubic Polynomial Trendline to a Scatter Chart and Capture Its Equation with Aspose.Cells for .NET
// Description: This example demonstrates how to create a workbook, fill X/Y data, add a scatter chart, attach a third‑order (cubic) polynomial trendline, display the regression equation and R‑squared on the chart, save the file, reload it, and read back the trendline order and equation‑visibility flag using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | scatter chart | polynomial trendline | cubic trendline | order 3 trendline | display equation | R-squared | chart series | retrieve trendline properties | Excel automation
// Common Searches: Aspose.Cells add cubic trendline | How to display polynomial equation on chart using Aspose.Cells | Retrieve trendline settings after saving workbook | Set trendline order to 3 in Aspose.Cells | Show R squared on scatter chart Aspose.Cells | Capture trendline properties from saved Excel file
// Developer Intent: Add a third‑order polynomial trendline to a scatter chart series, show its equation and R‑squared on the chart, persist the workbook, and programmatically read back the trendline configuration.
// Use Cases: Generate scientific or engineering reports with cubic regression curves embedded directly in Excel charts. | Automate quality‑control dashboards where a cubic fit and its formula must be visible to end users. | Validate that saved workbooks retain the correct trendline order and equation display before distribution. | Extract trendline parameters for further statistical calculations outside of Excel. | Create client‑ready Excel files that include the regression equation as part of the visual analysis.
// AI Prompts: Write C# code using Aspose.Cells to add a third‑order polynomial trendline to a scatter chart series, enable equation and R‑squared display, set the line color to red, save the workbook, then reload it and output the trendline order and equation visibility flag. | Show how to load an existing Excel file with Aspose.Cells and retrieve the Order and DisplayEquation properties of the first trendline in the first chart series. | Provide an example that populates X and Y data, creates a scatter chart, applies a cubic trendline, displays its regression formula, and prints the captured trendline details after reopening the file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// This example demonstrates how to create a workbook, fill X/Y data, add a scatter chart, attach a third‑order (cubic) polynomial trendline, display the regression equation and R‑squared on the chart, save the file, reload it, and read back the trendline order and equation‑visibility flag using Aspose.Cells for C#.
class PolynomialTrendlineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample X and Y data for a scatter chart
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 10; i++)
        {
            sheet.Cells["A" + i].PutValue(i);          // X values
            sheet.Cells["B" + i].PutValue(i * i + 5); // Y values (quadratic with noise)
        }

        // Add a scatter chart
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the series data (Y values) and X values
        chart.NSeries.Add("B2:B10", true);
        chart.NSeries[0].XValues = "A2:A10";
        chart.NSeries[0].Name = "Sample Series";

        // Add a polynomial trendline to the first series
        int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Polynomial);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

        // Set the polynomial order to 3 (cubic)
        trendline.Order = 3;

        // Display the equation (and R‑squared) on the chart
        trendline.DisplayEquation = true;
        trendline.DisplayRSquared = true;
        trendline.Color = Color.Red; // optional visual cue

        // Save the workbook with the chart
        string filePath = "PolynomialTrendline.xlsx";
        workbook.Save(filePath);

        // Reload the workbook to demonstrate capturing the trendline settings
        Workbook loadedWb = new Workbook(filePath);
        Chart loadedChart = loadedWb.Worksheets[0].Charts[0];
        Trendline loadedTrendline = loadedChart.NSeries[0].TrendLines[0];

        // Capture relevant information (order and that the equation is displayed)
        int capturedOrder = loadedTrendline.Order;
        bool equationDisplayed = loadedTrendline.DisplayEquation;

        // Output the captured details
        Console.WriteLine($"Trendline Order: {capturedOrder}");
        Console.WriteLine($"Equation displayed on chart: {equationDisplayed}");
    }
}
