// Title: Aspose.Cells for .NET – Create a Waterfall Chart with Custom Increase, Decrease and Total Colors (C#)
// Description: This example shows how to build a workbook, populate category/value data, add a Waterfall chart, and use reflection to detect each point’s role (total, increase, decrease). It then assigns Gold to total columns, LightGreen to increases, and LightCoral to decreases before saving the file as WaterfallCustomColors.xlsx.
// Keywords: Aspose.Cells waterfall chart C# | custom colors waterfall chart .NET | increase decrease total column color | WaterfallChartPoint reflection | Aspose.Cells chart point formatting | Excel waterfall chart customization
// Common Searches: change colors of increase and decrease columns in Aspose.Cells waterfall chart | use reflection to access WaterfallChartPoint properties C# | set custom total column color Aspose.Cells | waterfall chart custom styling Aspose.Cells .NET | how to format waterfall chart points programmatically
// Developer Intent: Generate a waterfall chart in an Excel workbook and apply distinct foreground and border colors to increase, decrease, and total columns programmatically.
// Use Cases: Financial reporting: highlight starting/ending totals in gold while showing revenue gains in green and cost losses in red. | Automated Excel generation: create workbooks at runtime with waterfall charts that use custom colors without referencing version‑specific types. | Dashboard export: produce printable Excel files where each waterfall segment is visually differentiated for quick stakeholder insight.
// AI Prompts: Rewrite the sample to use the strongly‑typed WaterfallChartPoint class instead of reflection for setting colors. | Add data labels with a custom number format to the waterfall chart while keeping the custom point colors. | Show how to apply a gradient fill to increase columns and a pattern fill to decrease columns in an Aspose.Cells waterfall chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Reflection;

// This example shows how to build a workbook, populate category/value data, add a Waterfall chart, and use reflection to detect each point’s role (total, increase, decrease). It then assigns Gold to total columns, LightGreen to increases, and LightCoral to decreases before saving the file as WaterfallCustomColors.xlsx.
class WaterfallChartWithCustomColors
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a waterfall chart
            // Column A – Categories, Column B – Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(0);          // total (starting point)

            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["B3"].PutValue(120);        // increase

            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["B4"].PutValue(-70);        // decrease

            sheet.Cells["A5"].PutValue("Profit");
            sheet.Cells["B5"].PutValue(0);          // total (ending point)

            // Add a Waterfall chart (type = Waterfall) to the worksheet
            // Parameters: topRow, leftColumn, bottomRow, rightColumn define the chart position
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including headers)
            chart.SetChartDataRange("A1:B5", true);

            // Access the first (and only) series of the waterfall chart
            Series series = chart.NSeries[0];

            // Iterate through each point and assign custom colors based on its role
            foreach (ChartPoint point in series.Points)
            {
                // Use reflection to access Waterfall-specific properties without requiring the WaterfallChartPoint type
                Type ptType = point.GetType();
                PropertyInfo isTotalProp = ptType.GetProperty("IsTotal");
                PropertyInfo isIncreaseProp = ptType.GetProperty("IsIncrease");

                bool isTotal = isTotalProp != null && (bool)isTotalProp.GetValue(point);
                bool isIncrease = isIncreaseProp != null && (bool)isIncreaseProp.GetValue(point);

                if (isTotal) // Total columns (start or end)
                {
                    point.Area.ForegroundColor = Color.Gold;
                    point.Border.Color = Color.DarkGoldenrod;
                }
                else if (isIncrease) // Increase columns
                {
                    point.Area.ForegroundColor = Color.LightGreen;
                    point.Border.Color = Color.Green;
                }
                else // Decrease columns
                {
                    point.Area.ForegroundColor = Color.LightCoral;
                    point.Border.Color = Color.Red;
                }
            }

            // Optional: force chart layout calculation before saving
            chart.Calculate();

            // Save the workbook with the customized waterfall chart
            string outputPath = "WaterfallCustomColors.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
