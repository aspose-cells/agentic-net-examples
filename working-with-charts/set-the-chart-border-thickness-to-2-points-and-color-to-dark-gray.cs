// Title: Aspose.Cells for .NET – Set Chart Area Border to 2 pt Dark Gray (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, and configures the chart's ChartArea.Border with a 2‑point thickness and dark‑gray color before saving as ChartBorderDemo.xlsx.
// Keywords: Aspose.Cells chart border | C# set chart area thickness | dark gray chart border Aspose | chart styling Aspose.Cells .NET | ChartArea.Border WeightPt
// Common Searches: how to change chart border thickness in Aspose.Cells C# | set chart area color to dark gray using Aspose.Cells | Aspose.Cells example for customizing chart borders | C# Aspose.Cells chart border weight
// Developer Intent: Apply a 2‑point dark‑gray border to a chart area in an Excel file using Aspose.Cells for .NET.
// Use Cases: Standardize chart appearance in automated financial reporting. | Maintain a consistent visual style for charts across multiple workbooks. | Improve chart readability in presentations by adding a uniform dark‑gray border.
// AI Prompts: Write C# code that sets a chart's border to 3 pt red with Aspose.Cells. | Show how to assign different border colors and thicknesses to several charts on the same worksheet using Aspose.Cells for .NET. | Provide an example that customizes both ChartArea and PlotArea borders in a line chart with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, and configures the chart's ChartArea.Border with a 2‑point thickness and dark‑gray color before saving as ChartBorderDemo.xlsx.
    public class ChartBorderDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set chart area border
                Line chartBorder = chart.ChartArea.Border;
                chartBorder.Color = Color.DarkGray;   // Border color
                chartBorder.WeightPt = 2.0;           // Border thickness

                // Save the workbook
                workbook.Save("ChartBorderDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ChartBorderDemo.Run();
        }
    }
}
