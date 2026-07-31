// Title: Aspose.Cells C# – Set custom data labels from a separate column in a column chart
// Description: Creates a workbook, fills column A with categories, column B with values, and column C with custom label text. Adds a column chart, links series values (B2:B4) and X‑axis categories (A2:A4), then enables ShowCellRange and sets LinkedSource to C2:C4 so the chart displays those cells as data labels while hiding the numeric value and default category name. Saves the result as an XLSX file.
// Keywords: Aspose.Cells custom data labels | C# chart data labels from cell range | ShowCellRange Aspose.Cells | LinkedSource property .NET | column chart custom labels | hide chart values Aspose.Cells | Excel chart label customization | Aspose.Cells example C#
// Common Searches: Aspose.Cells set data labels from another column | C# chart custom labels using LinkedSource | How to hide values in Aspose.Cells chart labels | Show cell range as data labels Aspose.Cells | Create column chart with custom text labels .NET
// Developer Intent: Configure a column chart so its data labels are taken from a dedicated worksheet column instead of the default value or category text.
// Use Cases: Display descriptive names (e.g., Alpha, Beta, Gamma) on each bar of a column chart for clearer presentation. | Build financial or sales dashboards where label text is stored in cells, allowing non‑developers to edit labels without code changes. | Replace numeric data labels with custom identifiers to improve readability in exported Excel reports.
// AI Prompts: Generate C# code with Aspose.Cells that links a chart's data labels to a cell range and suppresses the numeric values. | Explain how ShowCellRange and LinkedSource affect chart labeling in Aspose.Cells for .NET. | Show an example of assigning different label ranges to multiple series in a single Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // Creates a workbook, fills column A with categories, column B with values, and column C with custom label text. Adds a column chart, links series values (B2:B4) and X‑axis categories (A2:A4), then enables ShowCellRange and sets LinkedSource to C2:C4 so the chart displays those cells as data labels while hiding the numeric value and default category name. Saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data:
            // Column A – categories (used for X‑axis)
            // Column B – numeric values (series data)
            // Column C – custom labels that will be shown as data labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("Alpha");
            sheet.Cells["C3"].PutValue("Beta");
            sheet.Cells["C4"].PutValue("Gamma");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data (values) and the category (X‑axis) data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels to use the custom label column (C2:C4)
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowCellRange = true;          // Enable using a cell range for labels
            dataLabels.LinkedSource = "C2:C4";        // Reference to the custom label cells
            dataLabels.ShowValue = false;            // Hide the numeric value
            dataLabels.ShowCategoryName = false;     // Hide the default category name

            // Save the workbook
            workbook.Save("ChartWithCustomCategoryLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}
