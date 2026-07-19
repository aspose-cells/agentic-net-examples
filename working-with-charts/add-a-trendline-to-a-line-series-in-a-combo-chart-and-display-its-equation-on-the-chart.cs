// Title: Create a Combo Chart with Trendline Equation in Aspose.Cells (C#)
// Description: C# sample that builds an Excel workbook, fills month, sales and profit data, adds a combo chart (column + line), applies a linear trendline to the profit line, shows the formula on the chart, customizes the line color, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells combo chart | C# trendline | Excel line series trendline | display trendline equation | customize trendline color | hide R-squared Aspose | chart series conversion | automated Excel reporting
// Common Searches: Aspose.Cells add trendline to combo chart | show equation of trendline in Aspose.Cells C# | change trendline color Aspose.Cells chart | hide R squared Aspose.Cells trendline | convert column series to line in combo chart Aspose
// Developer Intent: Insert a linear trendline into the line series of a mixed chart and render its formula directly on the chart with Aspose.Cells for .NET.
// Use Cases: Financial report that combines sales columns with a profit line and includes a forecast equation. | Dashboard workbook where trend analysis is visualized on a mixed chart for executive review. | Automated Excel generation for presentations that require custom trendline styling and equation display.
// AI Prompts: Generate C# code using Aspose.Cells to create a combo chart (column + line) and add a linear trendline that shows its formula. | Show how to set the trendline color to blue and suppress the R‑squared value in an Aspose.Cells chart. | Explain how to retrieve the calculated trendline equation from a combo chart via the Aspose.Cells API.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsComboTrendlineDemo
{
    // C# sample that builds an Excel workbook, fills month, sales and profit data, adds a combo chart (column + line), applies a linear trendline to the profit line, shows the formula on the chart, customizes the line color, and saves the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Categories (e.g., months)
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

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

                // Add a Combo chart (initially a Column chart)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // First series: Column (Sales)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Sales";

                // Second series: Line (Profit)
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Profit";
                chart.NSeries[1].Type = ChartType.Line; // Convert second series to line for combo effect

                // Note: Category (X) axis data is automatically taken from the first column (A2:A5)

                // Add a linear trendline to the line series (index 1)
                int trendlineIdx = chart.NSeries[1].TrendLines.Add(TrendlineType.Linear);
                Trendline trendline = chart.NSeries[1].TrendLines[trendlineIdx];

                // Display the equation on the chart
                trendline.DisplayEquation = true;
                // Optionally hide R-squared value
                trendline.DisplayRSquared = false;
                // Customize appearance (optional)
                trendline.Color = Color.Blue;

                // Save the workbook
                string outputPath = "ComboChartWithTrendline.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
