// Title: Create a Column Chart, Change Its Title, and Save as ChartReport.xlsx Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to generate a new workbook with sample data, add a column chart, modify the chart title, and export the workbook to a new XLSX file named ChartReport.xlsx using Aspose.Cells in C#.
// Keywords: Aspose.Cells C# chart example | add column chart Aspose.Cells | update chart title programmatically | save workbook as XLSX Aspose.Cells | export chart report .NET | Aspose.Cells chart manipulation | C# Excel chart automation
// Common Searches: how to add a column chart with Aspose.Cells .NET | change chart title using Aspose.Cells C# | save Excel file with chart Aspose.Cells | Aspose.Cells example ChartReport.xlsx | C# code to create and export chart in Excel
// Developer Intent: Generate an Excel file, insert a column chart, customize its title, and write the result to ChartReport.xlsx.
// Use Cases: Automated sales dashboards that embed a column chart and are delivered as a ready‑to‑share XLSX file. | Monthly performance reports where the chart title reflects the reporting period before the workbook is archived. | Utility tools that let users rename chart titles on the fly and export the updated workbook for downstream processing.
// AI Prompts: Write C# code with Aspose.Cells to add a line chart from range D2:D15, set axis labels, and save as LineReport.xlsx. | Show how to programmatically adjust legend position, axis titles, and data labels for a chart, then save the workbook using Aspose.Cells. | Explain how to export a chart as a PNG image while also saving the containing workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartReport
{
    // Demonstrates how to generate a new workbook with sample data, add a column chart, modify the chart title, and export the workbook to a new XLSX file named ChartReport.xlsx using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
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

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Example chart update: change the title
            chart.Title.Text = "Updated Chart Title";

            // Save the modified workbook to a new file named ChartReport.xlsx
            workbook.Save("ChartReport.xlsx", SaveFormat.Xlsx);
        }
    }
}
