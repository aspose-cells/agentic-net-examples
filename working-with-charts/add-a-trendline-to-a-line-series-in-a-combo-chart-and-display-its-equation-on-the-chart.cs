// Title: Add a Linear Trendline with Equation to a Line Series in a Combo Chart using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, fills it with quarterly sales and profit data, inserts a column‑line combo chart, converts the profit series to a line, adds a linear trendline, and displays the equation directly on the chart before saving the file as XLSX.
// Keywords: Aspose.Cells C# combo chart | add trendline to Excel chart .NET | display trendline equation | linear trendline Aspose.Cells | column and line series chart | Excel automation C# | GitHub Aspose.Cells example
// Common Searches: how to add a trendline to a line series in a combo chart with Aspose.Cells | show equation of a trendline on an Excel chart using C# | Aspose.Cells combo chart with column and line series | C# code to create a trendline in an Excel workbook
// Developer Intent: Insert a linear trendline on the line series of a combo chart and make its equation visible.
// Use Cases: Quarterly financial reports that need a profit trend equation for forecasting. | Dashboards that combine sales columns with profit lines and highlight trend analysis. | Automated Excel generation for data‑driven presentations that require trendline annotations.
// AI Prompts: Generate C# code to add a polynomial trendline to the second series of a combo chart and show both the equation and R‑squared value. | Explain how to customize the trendline’s line style, color, and thickness in the provided Aspose.Cells example. | Provide a method to extract the calculated trendline equation from the chart object after saving the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// This example creates a workbook, fills it with quarterly sales and profit data, inserts a column‑line combo chart, converts the profit series to a line, adds a linear trendline, and displays the equation directly on the chart before saving the file as XLSX.
class AddTrendlineToComboChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Column series data (e.g., Sales)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(170);
            sheet.Cells["B5"].PutValue(200);

            // Line series data (e.g., Profit)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(55);
            sheet.Cells["C5"].PutValue(70);

            // Add a combo chart (Column + Line). Use a Column chart as base and set second series type to Line.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // First series – column (Sales)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Sales";

            // Second series – line (Profit)
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].Name = "Profit";
            chart.NSeries[1].Type = ChartType.Line; // Plot as line

            // Add a linear trendline to the line series (index 1)
            int trendlineIdx = chart.NSeries[1].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[1].TrendLines[trendlineIdx];
            trendline.DisplayEquation = true;   // Show equation on the chart
            trendline.DisplayRSquared = false; // Hide R‑squared (optional)

            // Prepare output path and ensure directory exists
            string outputPath = "ComboChartWithTrendline.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
