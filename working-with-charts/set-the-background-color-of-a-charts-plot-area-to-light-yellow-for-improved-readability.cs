// Title: Aspose.Cells for .NET – Set Chart Plot Area Background to Light Yellow (C#)
// Description: Creates a new workbook, adds a column chart from sample data, and changes the chart's plot area background to LightYellow using the Aspose.Cells API (chart.PlotArea.Area.BackgroundColor). The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells | C# chart background color | plot area background Aspose.Cells | light yellow chart plot area | Chart.PlotArea.Area.BackgroundColor | Excel chart styling .NET | set chart plot area color | Aspose.Cells chart formatting | Excel workbook chart example | Aspose.Cells API
// Common Searches: how to change chart plot area background color with Aspose.Cells C# | Aspose.Cells set plot area to light yellow | C# code for chart background color Aspose.Cells | Aspose.Cells chart formatting examples | change Excel chart plot area color using .NET
// Developer Intent: Apply a light‑yellow background to a chart’s plot area in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enhance readability of column charts by adding a subtle background shade. | Generate reports where chart plot areas need a consistent visual theme. | Automate styling of multiple charts across workbooks for corporate branding.
// AI Prompts: Write C# code with Aspose.Cells that sets the plot area background of any chart type to LightYellow and saves the file. | Show how to apply gradient or border styles to a chart’s plot area after setting the background color using Aspose.Cells. | Provide a step‑by‑step guide for customizing plot area colors for pie, line, and bar charts in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds a column chart from sample data, and changes the chart's plot area background to LightYellow using the Aspose.Cells API (chart.PlotArea.Area.BackgroundColor). The workbook is then saved as an XLSX file.
    public class SetPlotAreaBackgroundColor
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
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
                chart.SetChartDataRange("A1:B4", true);

                // Set the plot area background color to light yellow
                chart.PlotArea.Area.BackgroundColor = Color.LightYellow;

                // Save the workbook
                string outputPath = "ChartPlotAreaLightYellow.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while creating the chart: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                SetPlotAreaBackgroundColor.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
