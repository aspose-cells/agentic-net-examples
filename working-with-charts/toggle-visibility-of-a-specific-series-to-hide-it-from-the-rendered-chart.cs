// Title: Hide a chart series in Aspose.Cells (C#) with the IsFiltered property
// Description: Creates a workbook, adds two data series, builds a column chart, assigns ranges, then sets the first series' IsFiltered flag to true to exclude it from the rendered chart before saving the file.
// Keywords: Aspose.Cells | C# | chart series visibility | IsFiltered | column chart | hide series | toggle series | Excel automation | chart filtering | Aspose.Cells API
// Common Searches: Aspose.Cells hide chart series C# | IsFiltered property example Aspose.Cells | how to hide a data series in Aspose chart | toggle series visibility programmatically Aspose.Cells | remove series from chart without deleting data Aspose
// Developer Intent: Programmatically prevent a selected data series from appearing in a generated Excel chart while retaining its underlying data.
// Use Cases: Provide drill‑down reports where optional series can be shown on demand. | Create chart templates with optional series that users can enable or disable. | Exclude outlier or confidential data from visualizations while keeping it in the workbook. | Generate dynamic dashboards that hide series based on runtime conditions.
// AI Prompts: Give C# code to hide a series in an Aspose.Cells chart using IsFiltered. | Show how to toggle visibility of multiple series in an Aspose.Cells chart based on a condition. | Explain the effect of IsFiltered = true versus removing a series from the NSeries collection. | Provide a sample that restores a hidden series in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideSeriesDemo
{
    // Creates a workbook, adds two data series, builds a column chart, assigns ranges, then sets the first series' IsFiltered flag to true to exclude it from the rendered chart before saving the file.
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

            // Set data ranges for the two series
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the first series (index 0) using IsFiltered property
            chart.NSeries[0].IsFiltered = true;

            // Save the workbook
            workbook.Save("HideSeriesDemo.xlsx");
        }
    }
}
