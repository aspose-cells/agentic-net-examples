// Title: Add a thick solid border to a chart area with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example for adding a wide solid border around a chart | how to set the chart area border weight to wide line in Aspose.Cells .NET | apply a black solid line border to an Excel column chart using Aspose.Cells | Aspose.Cells chart outer border customization tutorial for .xlsx files | set solid dash style for chart area border with Aspose.Cells in C#
// Tags: Aspose.Cells chart border styling | Aspose.Cells chart line weight setting | solid dash border for Excel chart .NET | column chart outer border Aspose.Cells | C# Excel chart border customization

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartBorderDemo
{
    // The sample creates a workbook, inserts a column chart with sample data, and uses the ChartArea.Border properties (Weight, DashType, Color) to apply a thick, solid, black border around the chart before saving the file as an .xlsx workbook.
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
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");
                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(150);
                worksheet.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Optional chart title
                chart.Title.Text = "Quarterly Sales";

                // ------------------------------
                // Add a thick solid border around the entire chart
                // ------------------------------

                // Set outer border of the chart area
                chart.ChartArea.Border.Weight = WeightType.WideLine;   // Thick line (wide)
                chart.ChartArea.Border.DashType = MsoLineDashStyle.Solid; // Solid dash style
                chart.ChartArea.Border.Color = Color.Black;            // Black color

                // Uncomment to add a border to the plot area as well
                // chart.PlotArea.Border.Weight = WeightType.WideLine;
                // chart.PlotArea.Border.DashType = MsoLineDashStyle.Solid;
                // chart.PlotArea.Border.Color = Color.DarkGray;

                // Save the workbook
                string outputPath = "ChartWithThickBorder.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The chart now has a thick solid border around it.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
