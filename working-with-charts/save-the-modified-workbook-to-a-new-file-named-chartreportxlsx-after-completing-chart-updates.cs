// Title: Create a Column Chart and Save as ChartReport.xlsx with Aspose.Cells for .NET
// Description: Demonstrates how to build a new workbook, populate cells A1:B4, add a column chart with a "Sales Overview" title, and save the file as ChartReport.xlsx using Aspose.Cells in C#.
// Keywords: Aspose.Cells column chart C# | save workbook as xlsx Aspose | add chart to worksheet .NET | set chart title programmatically | generate Excel chart report
// Common Searches: Aspose.Cells add column chart example | C# save workbook with chart Aspose | how to set chart title Aspose.Cells | export chart to ChartReport.xlsx | create Excel chart without Excel UI
// Developer Intent: Create an Excel file, insert a column chart with data and a custom title, then write the workbook to ChartReport.xlsx.
// Use Cases: Automate sales‑summary charts for daily email distribution. | Generate chart‑rich Excel reports in a server‑side .NET service. | Produce workbook templates with embedded visualizations for downstream analysis.
// AI Prompts: Generate C# code that adds a line chart to an existing Aspose.Cells workbook, sets axis labels, and saves it as Report.xlsx. | Explain how to change a column chart to a pie chart, modify the legend, and export the workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartReport
{
    // Demonstrates how to build a new workbook, populate cells A1:B4, add a column chart with a "Sales Overview" title, and save the file as ChartReport.xlsx using Aspose.Cells in C#.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
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

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Example chart update: change the chart title
            chart.Title.Text = "Sales Overview";

            // Save the modified workbook to a new file named ChartReport.xlsx
            workbook.Save("ChartReport.xlsx", SaveFormat.Xlsx);
        }
    }
}
