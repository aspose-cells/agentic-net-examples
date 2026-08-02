// Title: Add a 3‑D Column Chart to an Aspose.Cells Worksheet (C#)
// Description: Creates a new workbook, populates cells A1:B5 with quarterly sales figures, inserts a 3‑D column chart (ChartType.Column3D) positioned rows 6‑20 and columns A‑J, binds the range A1:B5 (including headers), adds a title and legend, and saves the file as 3DColumnChart.xlsx.
// Keywords: Aspose.Cells | C# | 3D column chart | ChartType.Column3D | add chart to worksheet | set chart data range | chart title | chart legend | export XLSX | Aspose.Cells example
// Common Searches: Aspose.Cells add 3D column chart C# | C# Aspose.Cells set chart data range | How to add title and legend to a 3D column chart in Aspose.Cells | Insert chart into worksheet Aspose.Cells .NET | ChartType.Column3D usage example
// Developer Intent: Insert and configure a three‑dimensional column chart in a worksheet using Aspose.Cells for .NET.
// Use Cases: Generate a sales report workbook that visualizes quarterly figures with a 3‑D column chart. | Build a performance dashboard that places multiple 3‑D column charts for side‑by‑side metric comparison. | Automate periodic financial exports that include titled charts and legends for stakeholder presentations.
// AI Prompts: Write C# code with Aspose.Cells to add a 3‑D column chart from a dynamic range, customize its colors, and save the workbook as PDF. | Provide a step‑by‑step tutorial for inserting a 3‑D column chart, binding data, adding a title and legend, and exporting the result as XLSX using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCells3DColumnChart
{
    // Creates a new workbook, populates cells A1:B5 with quarterly sales figures, inserts a 3‑D column chart (ChartType.Column3D) positioned rows 6‑20 and columns A‑J, binds the range A1:B5 (including headers), adds a title and legend, and saves the file as 3DColumnChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["A5"].PutValue("Q4");

            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(1200);
            worksheet.Cells["B3"].PutValue(1500);
            worksheet.Cells["B4"].PutValue(1800);
            worksheet.Cells["B5"].PutValue(2100);

            // Add a three‑dimensional column chart (Column3D)
            // Parameters: chart type, top row, left column, bottom row, right column
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 6, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Define the data range for the chart (including headers)
            chart.SetChartDataRange("A1:B5", true);

            // Optional: set a title and enable the legend
            chart.Title.Text = "Quarterly Sales (3D)";
            chart.ShowLegend = true;

            // Save the workbook
            workbook.Save("3DColumnChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
