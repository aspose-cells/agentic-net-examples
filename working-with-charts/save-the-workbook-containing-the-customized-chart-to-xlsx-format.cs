// Title: Aspose.Cells C# Example: Save Workbook with a Customized Column Chart to XLSX
// Description: Demonstrates how to create a new Workbook, populate it with sample data, add a column chart, customize the title and legend, and persist the workbook—including the chart—to an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | save workbook as XLSX | column chart | chart customization | Export chart to Excel | SaveFormat.Xlsx | Aspose.Cells chart example | Excel automation .NET | programmatic Excel chart
// Common Searches: Aspose.Cells save workbook with chart C# | how to export a chart to XLSX using Aspose.Cells | C# create column chart and save Excel file | Aspose.Cells example save workbook containing chart | save Aspose.Cells workbook as .xlsx
// Developer Intent: Persist a workbook that contains a customized column chart as an XLSX file using Aspose.Cells for .NET.
// Use Cases: Generate a sales dashboard with a column chart and automatically deliver the .xlsx report to stakeholders. | Create inventory visualizations from database queries, style the chart, and store the result for downstream analysis. | Schedule a C# service to produce weekly performance charts and archive each workbook in XLSX format.
// AI Prompts: Write C# code with Aspose.Cells to add a line chart, set a custom title and legend position, then save the workbook as XLSX. | Explain how to modify the font size of a chart title and change the legend placement before saving the workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveExample
{
    // Demonstrates how to create a new Workbook, populate it with sample data, add a column chart, customize the title and legend, and persist the workbook—including the chart—to an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apples");
            worksheet.Cells["A3"].PutValue("Bananas");
            worksheet.Cells["A4"].PutValue("Cherries");

            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["B3"].PutValue(45);
            worksheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);               // Values
            chart.NSeries.CategoryData = "A2:A4";           // Categories

            // Customize the chart (example: set title and legend)
            chart.Title.Text = "Fruit Quantity";
            chart.Title.Font.Size = 14;
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save the workbook containing the customized chart to XLSX format
            workbook.Save("CustomizedChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
