// Title: How to set a chart’s border thickness to 2 points and color it dark gray with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart border weight to 2 points | Change chart area border color to dark gray using Aspose.Cells .NET | C# code to format chart border thickness in Excel with Aspose.Cells | Set chart area line style in an Aspose.Cells workbook | Example for customizing chart border color and weight in Aspose.Cells | Apply dark gray border to a column chart with Aspose.Cells C#
// Tags: chart area border weight Aspose.Cells | chart area border color Aspose.Cells | set chart border thickness .NET | customize chart border styling C# | Aspose.Cells chart formatting example | column chart border formatting Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for WeightType enum

namespace AsposeCellsChartBorderDemo
{
    // The sample creates a new workbook, adds a column chart with sample data, and uses the ChartArea.Border property to set the border weight to WeightType.WideLine (2 points) and the border color to Color.DarkGray, then saves the workbook as ChartBorderDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["B3"].PutValue(45);
                worksheet.Cells["B4"].PutValue(25);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // ---- Set chart border thickness and color ----
                // Use ChartArea.Border (Line) to define the outer border of the chart
                chart.ChartArea.Border.Weight = WeightType.WideLine;   // 2 points thickness
                chart.ChartArea.Border.Color = Color.DarkGray;        // Dark gray color

                // Optional: give the chart a title to verify it was created
                chart.Title.Text = "Sample Column Chart";

                // Save the workbook to a file
                string outputPath = "ChartBorderDemo.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Chart border thickness set to 2 points and color set to DarkGray.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
