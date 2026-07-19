// Title: Toggle visibility of a chart series in Aspose.Cells (C#) using the IsFiltered property
// Description: Creates a workbook, adds two data series, builds a column chart, then hides the second series by setting its IsFiltered flag to true before saving the file.
// Keywords: Aspose.Cells chart series visibility | C# IsFiltered property | hide chart series Aspose.Cells | toggle series visibility .NET | Aspose.Cells column chart example
// Common Searches: Aspose.Cells hide series C# | IsFiltered chart series Aspose.Cells example | how to hide a series in a column chart using Aspose.Cells | toggle chart series visibility programmatically Aspose
// Developer Intent: Programmatically exclude a selected data series from rendering in an Aspose.Cells chart while preserving the chart structure.
// Use Cases: Provide drill‑down reports where optional series can be shown or hidden on demand. | Create reusable chart templates with toggleable series for user‑driven analysis. | Generate multi‑series charts but suppress outlier or confidential series before distribution.
// AI Prompts: Show C# code that hides a specific series in an Aspose.Cells chart using IsFiltered. | Give an example of toggling visibility for multiple chart series in Aspose.Cells based on a condition. | Explain the impact of the IsFiltered property on chart rendering in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesVisibilityDemo
{
    // Creates a workbook, adds two data series, builds a column chart, then hides the second series by setting its IsFiltered flag to true before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add both series to the chart
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the second series (index 1) by setting IsFiltered to true
            chart.NSeries[1].IsFiltered = true;

            // Save the workbook
            workbook.Save("SeriesVisibilityDemo.xlsx");
        }
    }
}
