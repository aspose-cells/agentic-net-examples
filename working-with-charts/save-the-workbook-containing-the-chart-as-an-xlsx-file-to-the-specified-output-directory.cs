// Title: Save an Aspose.Cells Workbook with a Column Chart to a Specified Folder (C#)
// Description: Creates a new Workbook, populates sample data, adds a column chart, builds a full file path from a given output directory, ensures the folder exists, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# save workbook | export chart to XLSX | column chart Aspose.Cells | create Excel file with chart | save workbook to custom directory | Aspose.Cells file path handling | C# generate charted workbook | Aspose.Cells SaveFormat.Xlsx | directory creation C# Aspose
// Common Searches: how to save an Aspose.Cells workbook with a chart to a specific folder in C# | Aspose.Cells create column chart and export as XLSX | C# save Excel file with chart to user‑defined directory | Aspose.Cells workbook file path example | save charted workbook to network share using Aspose.Cells
// Developer Intent: Save a chart‑containing workbook as an XLSX file in a user‑specified output directory.
// Use Cases: Generate a sales dashboard with a column chart and store the Excel file in a client‑provided folder. | Automate nightly report creation on a server and write the charted workbook to a shared network location. | Create a temporary workbook with a chart, save it to a temp directory, and attach it to an email from a background service.
// AI Prompts: Write C# code that builds an Aspose.Cells workbook with a pie chart and saves it to a path supplied at runtime. | Explain how to add a chart title, axis labels, and legend to the column chart before saving the workbook. | Show how to modify the example to export the same workbook as a PDF while preserving the chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveExample
{
    // Creates a new Workbook, populates sample data, adds a column chart, builds a full file path from a given output directory, ensures the folder exists, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
    public class ChartWorkbookSaver
    {
        /// <param name="outputDirectory">Full path of the directory where the file will be saved.</param>
        public static void Run(string outputDirectory)
        {
            try
            {
                // Ensure the output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["A4"].PutValue("Cherries");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Build the full file path
                string filePath = Path.Combine(outputDirectory, "ChartWorkbook.xlsx");

                // Save the workbook as XLSX
                workbook.Save(filePath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook with chart saved to: {filePath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            // Determine output directory: use first argument or current directory
            string outputDir = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();
            ChartWorkbookSaver.Run(outputDir);
        }
    }
}
