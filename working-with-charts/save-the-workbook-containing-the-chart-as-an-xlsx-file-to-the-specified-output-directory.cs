// Title: Save a C# Aspose.Cells Workbook with a Column Chart to a Specified Folder (XLSX)
// Description: Creates a new Workbook, adds sample data, inserts a column chart linked to that data, builds a file path from a user‑provided output directory, and saves the workbook (including the chart) as an XLSX file using Aspose.Cells.
// Keywords: Aspose.Cells C# save workbook | export chart to XLSX | column chart Aspose.Cells example | custom output directory Aspose.Cells | Workbook.Save with chart | Aspose.Cells SaveFormat.Xlsx | C# chart workbook export
// Common Searches: C# Aspose.Cells save workbook with chart to folder | How to export a charted spreadsheet using Aspose.Cells | Aspose.Cells save XLSX to specific path | Create column chart and save workbook Aspose.Cells C# | Aspose.Cells example save chart to network drive
// Developer Intent: Generate a workbook that contains a column chart and write it as an XLSX file to a user‑defined directory.
// Use Cases: Produce a sales report with a column chart and automatically store the file in a shared network folder. | Run a batch job that creates multiple charted spreadsheets, each saved to a project‑specific output directory. | Provide a console utility where end users specify the destination path for the chart‑enabled workbook.
// AI Prompts: Write C# code that uses Aspose.Cells to add a pie chart to a workbook and saves the file as XLSX to a user‑provided directory. | Show how to modify this example to create a line chart instead of a column chart and save it to a configurable path. | Provide a method that accepts data arrays, builds a bar chart with Aspose.Cells, and returns the full file path of the saved XLSX workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new Workbook, adds sample data, inserts a column chart linked to that data, builds a file path from a user‑provided output directory, and saves the workbook (including the chart) as an XLSX file using Aspose.Cells.
class SaveChartWorkbook
{
    public static void Run(string outputDirectory)
    {
        try
        {
            // Ensure the output directory exists
            if (string.IsNullOrWhiteSpace(outputDirectory))
                outputDirectory = Directory.GetCurrentDirectory();

            Directory.CreateDirectory(outputDirectory);

            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["B3"].PutValue(45);

            // Add a column chart and set its data source
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Build the full file path for the XLSX file
            string filePath = Path.Combine(outputDirectory, "ChartWorkbook.xlsx");

            // Save the workbook (which includes the chart) as an XLSX file
            workbook.Save(filePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {filePath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error creating chart workbook: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        // Determine output directory (first argument or current directory)
        string outputDir = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();
        SaveChartWorkbook.Run(outputDir);
    }
}
