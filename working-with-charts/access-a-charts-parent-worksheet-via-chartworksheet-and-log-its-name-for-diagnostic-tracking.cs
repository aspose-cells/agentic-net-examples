// Title: Aspose.Cells .NET – Retrieve a Chart’s Containing Worksheet (Chart.Worksheet)
// Description: This example builds a workbook, adds sample data, creates a column chart, and then uses the Chart.Worksheet property to output the chart’s owner sheet name and index for diagnostic logging before saving the file.
// Keywords: Aspose.Cells Chart.Worksheet | C# get chart parent sheet | Aspose.Cells diagnostic logging | retrieve chart worksheet name .NET | chart worksheet index Aspose | Aspose.Cells example GitHub | US developers Aspose.Cells | global spreadsheet automation
// Common Searches: Aspose.Cells how to find worksheet of a chart | Chart.Worksheet property usage example | log chart's sheet name in C# Aspose | get chart parent worksheet index Aspose.Cells | debug chart placement Aspose.Cells .NET
// Developer Intent: Identify the worksheet that hosts a specific chart and record its name and index to aid debugging or reporting workflows.
// Use Cases: Confirm that automatically generated charts are placed on the intended sheet during report creation. | Iterate over all charts in a workbook and capture each chart’s sheet name and position for troubleshooting layout problems. | Include sheet details in custom log files before exporting charts to PDF or image formats.
// AI Prompts: Generate C# code that loops through every chart in a workbook and prints the chart’s sheet name and index using Aspose.Cells. | Show how to safely check for a null Chart.Worksheet reference and log an appropriate warning message. | Provide a snippet that records chart‑owner worksheet information to a structured log (JSON or CSV) before saving the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example builds a workbook, adds sample data, creates a column chart, and then uses the Chart.Worksheet property to output the chart’s owner sheet name and index for diagnostic logging before saving the file.
    public class ChartWorksheetDiagnosticDemo
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a recognizable name
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Populate some sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Define the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the chart's parent worksheet via Chart.Worksheet and log its details
            Console.WriteLine("Chart belongs to worksheet: " + chart.Worksheet.Name);
            Console.WriteLine("Worksheet index: " + chart.Worksheet.Index);

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartWorksheetDiagnosticDemo_out.xlsx");

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}
