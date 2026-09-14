// Title: Save a workbook with a column chart as QuarterlyReport.xlsx in an output folder using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a worksheet, fills quarter and revenue data, adds a column chart linked to those ranges, and saves the file as QuarterlyReport.xlsx in a specified output directory with Aspose.Cells. | Generate a .NET snippet that ensures an "output" folder exists, builds a column chart from cells A2:A5 and B2:B5, and writes the workbook to the folder in XLSX format using Aspose.Cells.
// Common Searches: asp.net how to add a column chart to an Excel workbook and save it to a custom folder with Aspose.Cells | c# Aspose.Cells generate quarterly revenue chart and export as QuarterlyReport.xlsx | save workbook containing chart to output directory using Aspose.Cells for .NET | create Excel file with chart from range B2:B5 and A2:A5 in C#
// Tags: Aspose.Cells create column chart | Aspose.Cells save workbook as xlsx | Aspose.Cells chart data source range | C# generate quarterly revenue Excel | Aspose.Cells ensure output directory

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, populates cells A1:B5 with quarter and revenue values, adds a column chart that references those ranges, ensures an "output" folder exists, and saves the workbook as QuarterlyReport.xlsx in that folder using Aspose.Cells for .NET.
    public class SaveWorkbookWithChart
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Quarter");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["A5"].PutValue("Q4");

            worksheet.Cells["B1"].PutValue("Revenue");
            worksheet.Cells["B2"].PutValue(15000);
            worksheet.Cells["B3"].PutValue(20000);
            worksheet.Cells["B4"].PutValue(18000);
            worksheet.Cells["B5"].PutValue(22000);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Ensure the output directory exists
            string outputFolder = "output";
            Directory.CreateDirectory(outputFolder);

            // Save the workbook as an XLSX file named QuarterlyReport.xlsx
            string outputPath = Path.Combine(outputFolder, "QuarterlyReport.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook with chart saved to: {outputPath}");
        }
    }
}
