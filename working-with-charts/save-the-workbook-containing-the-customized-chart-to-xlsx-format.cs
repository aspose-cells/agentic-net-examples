// Title: Save a Workbook with a Customized Column Chart to XLSX using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate quarterly sales data, add a column chart, customize its title and legend, and save the file as an XLSX document with Aspose.Cells for C#.
// Keywords: Aspose.Cells save XLSX | C# column chart Aspose.Cells | customize chart title legend Aspose.Cells | export chart to Excel .NET | Aspose.Cells workbook with chart
// Common Searches: Aspose.Cells save workbook with chart to XLSX | C# create column chart and export to Excel | how to set chart title and legend in Aspose.Cells | save Aspose.Cells workbook as .xlsx file | programmatically add and save chart using Aspose.Cells
// Developer Intent: Save a workbook that contains a customized column chart as an XLSX file.
// Use Cases: Generate a quarterly sales report with a column chart and distribute it as an Excel file. | Automate the creation of performance charts in a scheduled .NET job and persist them for later editing. | Build a template workbook, inject a chart programmatically, and save it for downstream users.
// AI Prompts: Write C# code with Aspose.Cells to add a line chart, format axes, and save the workbook as XLSX. | Show how to export a workbook containing multiple charts into separate XLSX files using Aspose.Cells. | Explain the steps to change chart title font size and legend position before saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveDemo
{
    // Demonstrates how to create a workbook, populate quarterly sales data, add a column chart, customize its title and legend, and save the file as an XLSX document with Aspose.Cells for C#.
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
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["A5"].PutValue("Q4");

            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["B4"].PutValue(180);
            worksheet.Cells["B5"].PutValue(200);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B5", true);               // Values
            chart.NSeries.CategoryData = "A2:A5";           // Categories

            // Customize the chart (example: set title and legend position)
            chart.Title.Text = "Quarterly Sales";
            chart.Title.Font.Size = 14;
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save the workbook with the customized chart to XLSX format
            workbook.Save("CustomizedChart.xlsx", SaveFormat.Xlsx);

            // Optional: inform the user
            Console.WriteLine("Workbook with customized chart saved as CustomizedChart.xlsx");
        }
    }
}
