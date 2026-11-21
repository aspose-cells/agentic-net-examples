// Title: How to save an Excel workbook with an updated chart title to XLSX while preserving all chart controls using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook with charts to XLSX preserving chart objects | update chart title in an Excel file and export with Aspose.Cells .NET | preserve Excel chart controls when saving a modified workbook using Aspose.Cells | C# code to modify the first chart title and save the workbook as XLSX with Aspose.Cells | ensure output directory exists before saving a workbook that contains charts in Aspose.Cells
// Tags: update first chart title Aspose.Cells | save workbook as XLSX with chart objects | Aspose.Cells retain chart controls | C# ensure output folder before workbook save | Aspose.Cells SaveFormat.Xlsx usage

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing XLSX workbook, checks for charts on the first worksheet, updates the title of the first chart, creates the output directory if needed, and then saves the workbook back to XLSX using Aspose.Cells, which automatically retains all chart controls. It also includes basic error handling for missing files and other exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook that contains the charts
            Workbook workbook = new Workbook(inputPath);

            // Example modification: update the title of the first chart
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count > 0)
            {
                Chart chart = sheet.Charts[0];
                chart.Title.Text = "Updated Chart Title";
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to XLSX format.
            // Aspose.Cells preserves all chart controls automatically.
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
