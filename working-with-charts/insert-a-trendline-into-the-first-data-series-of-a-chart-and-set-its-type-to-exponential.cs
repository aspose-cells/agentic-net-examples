// Title: Aspose.Cells .NET: Add Exponential Trendline to First Chart Series (C#)
// Description: Creates a workbook, populates sample data, adds a line chart, and inserts an exponential trendline into the first series. The trendline shows its equation, R‑squared value, and is colored red before saving the file.
// Keywords: Aspose.Cells C# exponential trendline | add trendline to chart Aspose.Cells | set trendline type exponential .NET | display trendline equation Aspose | show R-squared Aspose.Cells chart | change trendline color Aspose.Cells | Aspose.Cells chart example GitHub | C# Excel chart trendline Aspose
// Common Searches: How to add an exponential trendline in Aspose.Cells C# | Aspose.Cells set trendline color and display equation | C# code for chart trendline with R-squared using Aspose.Cells | Aspose.Cells example for exponential trendline on line chart | GitHub Aspose.Cells trendline sample
// Developer Intent: Add and customize an exponential trendline for the first series of a line chart using Aspose.Cells for .NET.
// Use Cases: Create a sales forecast workbook that visualizes exponential growth with a red trendline and shows the equation and R‑squared value. | Generate a scientific chart illustrating bacterial growth, highlighting the exponential trendline for quick interpretation. | Automate reporting dashboards where trendline styling and metrics are required for stakeholder presentations.
// AI Prompts: Provide C# Aspose.Cells code to insert an exponential trendline into the first series of a line chart and display its equation and R‑squared. | Show how to change the trendline color to blue and hide the equation in an Aspose.Cells chart. | Explain how to retrieve the equation string of an exponential trendline after adding it with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Creates a workbook, populates sample data, adds a line chart, and inserts an exponential trendline into the first series. The trendline shows its equation, R‑squared value, and is colored red before saving the file.
class InsertExponentialTrendline
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["A3"].PutValue(3);
            sheet.Cells["A4"].PutValue(4);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["B3"].PutValue(6);
            sheet.Cells["B4"].PutValue(8);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 15, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the first series (Y values)
            chart.NSeries.Add("B1:B4", true);

            // Set the category (X) data for the series
            chart.NSeries[0].XValues = "A1:A4";

            // Insert an exponential trendline into the first data series
            int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

            // Display equation, R‑squared value, and set a color
            trendline.DisplayEquation = true;
            trendline.DisplayRSquared = true;
            trendline.Color = Color.Red;

            // Save the workbook
            workbook.Save("ExponentialTrendline.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
