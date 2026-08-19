// Title: C# – Save a Workbook with a Column Chart as QuarterlyReport.xlsx using Aspose.Cells
// Description: Creates a new workbook, fills quarter and revenue data, adds a column chart, sets the chart title, ensures an "output" folder exists, and saves the file as QuarterlyReport.xlsx (XLSX) with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# save chart workbook | save workbook with chart Aspose.Cells | C# column chart Excel export | Aspose.Cells SaveFormat.Xlsx example | output folder Excel file C# | Aspose.Cells chart to XLSX | programmatic Excel report generation
// Common Searches: how to save an Aspose.Cells workbook that contains a chart to a folder in C# | Aspose.Cells example create column chart and export to XLSX | C# code to generate quarterly report with chart and save as Excel file | Aspose.Cells SaveFormat.Xlsx with chart example | save Excel file with chart to specific directory using Aspose.Cells
// Developer Intent: Programmatically save a workbook that includes a column chart as QuarterlyReport.xlsx in an "output" directory.
// Use Cases: Automated generation of quarterly revenue reports with embedded charts. | Batch processing of Excel files that require chart visuals before distribution. | Integrating chart‑enabled workbook export into web APIs or background services.
// AI Prompts: Show a C# example that creates a line chart from data and saves the workbook as Report.xlsx in a "results" folder using Aspose.Cells. | Explain how to change the chart title font size and export the same workbook to PDF instead of XLSX. | Provide step‑by‑step instructions to add multiple data series to a chart and save each workbook with a unique filename.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveDemo
{
    // Creates a new workbook, fills quarter and revenue data, adds a column chart, sets the chart title, ensures an "output" folder exists, and saves the file as QuarterlyReport.xlsx (XLSX) with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Quarter");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            sheet.Cells["B1"].PutValue("Revenue");
            sheet.Cells["B2"].PutValue(15000);
            sheet.Cells["B3"].PutValue(20000);
            sheet.Cells["B4"].PutValue(18000);
            sheet.Cells["B5"].PutValue(22000);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B5", true);               // Values
            chart.NSeries.CategoryData = "A2:A5";           // Categories
            chart.Title.Text = "Quarterly Revenue";

            // Define output path (ensure the folder exists)
            string outputPath = System.IO.Path.Combine("output", "QuarterlyReport.xlsx");
            System.IO.Directory.CreateDirectory("output");

            // Save the workbook as XLSX using the Save(string, SaveFormat) overload
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook with chart saved to: {outputPath}");
        }
    }
}
