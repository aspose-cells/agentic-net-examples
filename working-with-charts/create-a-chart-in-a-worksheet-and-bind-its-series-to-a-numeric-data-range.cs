// Title: C# – Create a Column Chart and Bind Its Series to a Numeric Range with Aspose.Cells
// Description: Demonstrates how to generate a new workbook, fill columns with category labels and numeric values, add a column chart, link the chart's NSeries to the range B2:B10, assign category data from A2:A10, set a chart title, and save the file as an Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells column chart C# | bind chart series numeric range | NSeries.Add Aspose.Cells | set category axis data Aspose.Cells | export Excel workbook .NET
// Common Searches: Aspose.Cells add column chart programmatically | How to bind chart series to a range in C# | Set category labels for Aspose.Cells chart | Save workbook with chart using Aspose.Cells
// Developer Intent: Programmatically create a column chart, connect its data series to numeric cells, define category labels, and write the result to an Excel file.
// Use Cases: Automated sales‑by‑category visual reports | Monthly KPI dashboards generated from database exports | Dynamic performance charts for multi‑region data sets
// AI Prompts: Generate code to add a second data series to the same column chart in Aspose.Cells. | Show how to customize axis titles, data labels, and marker styles after binding data. | Provide an example of exporting the created chart as a PNG image with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to generate a new workbook, fill columns with category labels and numeric values, add a column chart, link the chart's NSeries to the range B2:B10, assign category data from A2:A10, set a chart title, and save the file as an Excel workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample numeric data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 10; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart (ChartCollection.Add rule)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 2, 25, 11);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the series to the numeric data range (Chart.NSeries.Add rule)
        chart.NSeries.Add("=Sheet1!$B$2:$B$10", true);
        // Set category axis data
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$10";

        // Set a title for the chart
        chart.Title.Text = "Sample Numeric Chart";

        // Save the workbook
        workbook.Save("NumericChart.xlsx");
    }
}
