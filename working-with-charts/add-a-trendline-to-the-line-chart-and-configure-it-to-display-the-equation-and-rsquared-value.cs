// Title: Add Linear Trendline with Equation & R‑Squared to a Line Chart using Aspose.Cells for .NET
// Description: Creates a workbook, populates X/Y data, inserts a line chart, adds a linear trendline to the first series, and configures the trendline to show its regression equation and R‑squared value before saving as an Excel file.
// Keywords: Aspose.Cells C# trendline | line chart regression Aspose | display equation R‑squared .NET | Excel chart trendline API | Aspose.Cells chart customization | C# Excel linear regression | US developers Aspose.Cells tutorial | European .NET chart examples
// Common Searches: Aspose.Cells add linear trendline to chart | show regression equation on Excel chart with Aspose | C# Aspose.Cells display R‑squared value | how to add trendline in Aspose.Cells .NET | chart regression line Aspose.Cells example
// Developer Intent: Insert a trendline into a line chart and enable its equation and R‑squared display.
// Use Cases: Sales analysis reports that include a best‑fit line with statistical details. | Scientific experiment visualizations showing trend equations for data sets. | Automated KPI dashboards where stakeholders need quick insight into trend strength.
// AI Prompts: Generate C# code with Aspose.Cells to add a polynomial trendline and show its equation and R‑squared. | Explain how to style a trendline (color, dash type, thickness) in an Aspose.Cells chart. | Provide steps to extract the equation string from a Trendline object after saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, populates X/Y data, inserts a line chart, adds a linear trendline to the first series, and configures the trendline to show its regression equation and R‑squared value before saving as an Excel file.
class AddTrendlineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (X values in column A, Y values in column B)
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells["A" + i].PutValue(i - 1);               // X: 1,2,3,4,5
            sheet.Cells["B" + i].PutValue((i - 1) * 2 + 3);    // Y: 5,7,9,11,13 (example)
        }

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the series data (Y values) and the category data (X values)
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries.CategoryData = "A2:A6";

        // Add a linear trendline to the first series
        int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

        // Configure the trendline to display its equation and R‑squared value
        trendline.DisplayEquation = true;
        trendline.DisplayRSquared = true;

        // Save the workbook to a file
        workbook.Save("LineChartWithTrendline.xlsx");
    }
}
