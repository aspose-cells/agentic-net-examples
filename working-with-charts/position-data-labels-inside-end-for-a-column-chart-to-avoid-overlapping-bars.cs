// Title: Aspose.Cells C# – Set Column Chart Data Labels to Inside End to Avoid Overlap
// Description: Learn how to create a workbook with a column chart, enable data labels, and position them at the InsideEnd of each column using Aspose.Cells for .NET. This placement keeps labels inside the column tops, eliminating overlap with neighboring bars, and saves the result as an XLSX file.
// Keywords: Aspose.Cells data label position | C# column chart InsideEnd | Excel chart label placement | prevent overlapping chart labels | .NET chart data labels | Aspose.Cells column chart example
// Common Searches: Aspose.Cells set data label position InsideEnd | C# column chart label inside end | avoid overlapping data labels in Excel chart Aspose | how to position chart labels in Aspose.Cells | column chart label placement .NET
// Developer Intent: Place column chart data labels at the InsideEnd so they stay within each column and do not cover adjacent bars.
// Use Cases: Create sales dashboards where column values are displayed inside the bars for a clean look. | Generate financial reports with multiple column charts that keep labels readable without overlapping. | Automate Excel workbook production for marketing analytics, ensuring label clarity in dense column charts.
// AI Prompts: Show C# code to set column chart data labels to InsideEnd using Aspose.Cells. | How can I adjust label positions for several series in an Aspose.Cells column chart? | Explain customizing the appearance of data labels after applying the InsideEnd position in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLabelPositionExample
{
    // Learn how to create a workbook with a column chart, enable data labels, and position them at the InsideEnd of each column using Aspose.Cells for .NET. This placement keeps labels inside the column tops, eliminating overlap with neighboring bars, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Position data labels inside the end of each column to avoid overlapping bars
            series.DataLabels.Position = LabelPositionType.InsideEnd;

            // Save the workbook to an XLSX file
            workbook.Save("ColumnChart_InsideEndLabels.xlsx");
        }
    }
}
