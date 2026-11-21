// Title: Generate a Waterfall chart with custom increase, decrease, and total column colors using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom colors for waterfall chart increase and decrease columns | Create waterfall chart with separate total column color using Aspose.Cells .NET | C# example applying green red blue colors to waterfall chart points in Excel with Aspose.Cells | How to change point colors in Aspose.Cells waterfall chart based on value | Aspose.Cells chart formatting waterfall series point colors C#
// Tags: waterfall chart custom colors Aspose.Cells | set increase decrease colors Aspose.Cells chart | total column color Aspose.Cells waterfall | C# Aspose.Cells chart point formatting | Excel waterfall chart styling Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills category and value data, adds a Waterfall chart, assigns X‑axis categories, defines green, red, and blue colors for increase, decrease, and total columns, iterates through each data point to apply the appropriate foreground color based on its value, and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate data for the waterfall chart
            // Column A: Category names
            // Column B: Corresponding values
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(100);

            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["B3"].PutValue(150);

            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["B4"].PutValue(-70);

            sheet.Cells["A5"].PutValue("Profit");
            sheet.Cells["B5"].PutValue(180);

            sheet.Cells["A6"].PutValue("Total");
            sheet.Cells["B6"].PutValue(180); // Total column (usually the last point)

            // -------------------------------------------------
            // Add a Waterfall chart to the worksheet
            // -------------------------------------------------
            // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 27, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set a title for the chart
            chart.Title.Text = "Waterfall Chart Example";

            // Add a single series for the waterfall data (Y values)
            int seriesIndex = chart.NSeries.Add("B2:B6", true);

            // Set the X‑axis (category) data using XValues (compatible with current API)
            chart.NSeries[seriesIndex].XValues = "A2:A6";

            // -------------------------------------------------
            // Define custom colors for increase, decrease, and total columns
            // -------------------------------------------------
            Color increaseColor = Color.Green;   // Positive change
            Color decreaseColor = Color.Red;     // Negative change
            Color totalColor    = Color.Blue;    // Total column

            // Apply the colors to each data point
            for (int i = 0; i < chart.NSeries[seriesIndex].Points.Count; i++)
            {
                // Retrieve the actual value from the worksheet (row i+2 because data starts at row 2)
                double pointValue = sheet.Cells[i + 1, 1].DoubleValue; // Column B (zero‑based index 1)

                if (i == chart.NSeries[seriesIndex].Points.Count - 1) // Last point is treated as Total
                {
                    chart.NSeries[seriesIndex].Points[i].Area.ForegroundColor = totalColor;
                }
                else if (pointValue >= 0) // Increase column
                {
                    chart.NSeries[seriesIndex].Points[i].Area.ForegroundColor = increaseColor;
                }
                else // Decrease column
                {
                    chart.NSeries[seriesIndex].Points[i].Area.ForegroundColor = decreaseColor;
                }
            }

            // -------------------------------------------------
            // Save the workbook to an Excel file
            // -------------------------------------------------
            string outputPath = "WaterfallChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
