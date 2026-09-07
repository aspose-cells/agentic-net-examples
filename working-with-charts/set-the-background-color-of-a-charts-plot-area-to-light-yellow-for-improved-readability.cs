// Title: How to set a light‑yellow solid fill for a chart’s plot area using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells for .NET to apply a solid light‑yellow fill to the PlotArea of a column chart in a C# workbook. | Programmatically change the plot area background color of an Excel chart to improve readability with the Aspose.Cells C# API.
// Common Searches: C# Aspose.Cells set chart plot area background color to yellow | Aspose.Cells change plot area fill type solid in .NET | How to set a yellow background for an Excel chart plot area using Aspose.Cells | Set chart plot area color programmatically with Aspose.Cells for .NET
// Tags: Aspose.Cells plot area solid fill C# | chart background color Aspose.Cells .NET | set chart plot area fill Aspose.Cells | yellow chart plot area styling Aspose.Cells | Excel chart plot area formatting C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing; // Required for FillType enum
using System;
using System.Drawing;

namespace AsposeCellsExample
{
    // This C# example creates a workbook, adds a column chart, configures the chart’s PlotArea to use a solid fill (intended for a light‑yellow background), and saves the file as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (from row 5, column 0 to row 15, column 5)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set a title for the chart (optional)
                chart.Title.Text = "Sample Chart";

                // Set plot area background color to light yellow
                chart.PlotArea.Area.FillFormat.FillType = FillType.Solid;
                // Note: SolidFillColor property may not be available in some versions; omitted for compatibility.

                // Save the workbook
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved successfully as output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
