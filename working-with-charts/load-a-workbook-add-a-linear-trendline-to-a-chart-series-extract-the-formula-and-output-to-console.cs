// Title: Add a Linear Trendline to an Aspose.Cells Chart and Output Its Equation (C#)
// Description: Creates a workbook, fills A1:B4 with data, inserts a line chart, adds a linear trendline to the first series, enables equation display, saves the file, computes the regression line manually, and writes the trendline formula to the console.
// Keywords: Aspose.Cells C# chart trendline | linear regression Aspose.Cells | extract trendline equation .NET | add trendline to chart series | display chart trendline formula | save workbook with chart | trendline slope intercept calculation
// Common Searches: how to add a linear trendline in Aspose.Cells C# | retrieve trendline equation from Aspose.Cells chart | Aspose.Cells chart regression line example | C# code to display trendline formula in Excel workbook | save chart with trendline using Aspose.Cells
// Developer Intent: Add a linear trendline to a chart series, obtain its regression equation, and print the formula while saving the workbook.
// Use Cases: Automated reporting that includes a regression line on a line chart. | Logging the slope and intercept of a trendline for downstream analytics. | Generating Excel files with pre‑formatted charts for client delivery.
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart, adds a linear trendline, shows the equation, and prints the formula. | Provide a method to calculate slope and intercept from worksheet cells and return the trendline equation without using CellsAI. | Explain how to access and display the Trendline object's equation after adding it to an Aspose.Cells chart series.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills A1:B4 with data, inserts a line chart, adds a linear trendline to the first series, enables equation display, saves the file, computes the regression line manually, and writes the trendline formula to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["A3"].PutValue(3);
            sheet.Cells["A4"].PutValue(4);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["B3"].PutValue(6);
            sheet.Cells["B4"].PutValue(8);

            // Add a line chart and bind the data range
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:B4", true);

            // Add a linear trendline to the first series and enable its equation display
            int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];
            trendline.DisplayEquation = true;

            // Save the workbook to a file
            string filePath = "TrendlineDemo.xlsx";
            workbook.Save(filePath);

            // Calculate linear regression manually (since CellsAI is not available)
            double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
            int n = 4;
            for (int i = 1; i <= n; i++)
            {
                double x = Convert.ToDouble(sheet.Cells[$"A{i}"].Value);
                double y = Convert.ToDouble(sheet.Cells[$"B{i}"].Value);
                sumX += x;
                sumY += y;
                sumXY += x * y;
                sumX2 += x * x;
            }

            double slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
            double intercept = (sumY - slope * sumX) / n;
            string equation = $"y = {slope:F4}x + {intercept:F4}";

            // Output the extracted formula to the console
            Console.WriteLine("Trendline formula: " + equation);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
