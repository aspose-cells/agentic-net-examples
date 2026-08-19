// Title: Hide Chart Legend in Aspose.Cells .NET (C#) – Minimalist Column Chart
// Description: This C# example creates a workbook, adds sample data, builds a column chart, and removes the legend by setting Chart.ShowLegend = false. The resulting file (ChartWithoutLegend.xlsx) demonstrates a clean, presentation‑ready chart layout.
// Keywords: Aspose.Cells hide legend C# | chart.ShowLegend false | remove chart legend Aspose.Cells .NET | minimalist Excel chart | Aspose.Cells chart formatting | C# Excel chart legend hide
// Common Searches: how to hide legend in Aspose.Cells chart C# | Aspose.Cells remove chart legend example | minimalist chart Aspose.Cells .NET | Aspose.Cells ShowLegend property usage
// Developer Intent: Need to suppress the legend in an Aspose.Cells generated chart to achieve a cleaner visual for slides or reports.
// Use Cases: Slide decks where series names are self‑explanatory | Compact dashboards without redundant legends | Printed reports that require uncluttered charts
// AI Prompts: Generate code to hide the legend for any Aspose.Cells chart type in C#. | Show how to toggle Chart.ShowLegend based on a runtime condition. | Explain additional formatting steps (title, axis labels) after removing the legend in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendHide
{
    // This C# example creates a workbook, adds sample data, builds a column chart, and removes the legend by setting Chart.ShowLegend = false. The resulting file (ChartWithoutLegend.xlsx) demonstrates a clean, presentation‑ready chart layout.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to achieve a minimalist layout
            chart.ShowLegend = false;

            // Save the workbook to a file
            workbook.Save("ChartWithoutLegend.xlsx");
        }
    }
}
