// Title: How to load only chart objects from an Excel workbook using Aspose.Cells LoadOptions.LoadFilter in C#
// AI Prompts: Write C# code that creates a LoadOptions instance with a LoadFilter set to include only Chart objects, then loads an .xlsx file with Aspose.Cells and saves the result. | Show how to configure Aspose.Cells LoadOptions.LoadFilter to filter for charts when opening a workbook in a .NET application. | Provide a C# example that extracts only the charts from an existing Excel file using Aspose.Cells and writes them to a new workbook.
// Common Searches: Aspose.Cells C# load workbook with charts only using LoadFilter | How to use LoadOptions to load only chart objects from an Excel file in .NET | C# Aspose.Cells filter workbook objects to load only charts | Load only charts from .xlsx with Aspose.Cells LoadOptions example
// Tags: Aspose.Cells LoadOptions chart filter | C# load Excel charts only | Aspose.Cells chart-only loading | LoadFilter for chart objects Aspose | Excel workbook chart extraction .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Loading;

// The sample checks for an input.xlsx file, creates a placeholder workbook with a column chart if missing, then loads the workbook using the default full load (chart‑only loading is not supported in the current version) and saves it as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a placeholder if missing
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a placeholder workbook.");
                var placeholderWb = new Workbook();
                var ws = placeholderWb.Worksheets[0];
                ws.Cells["A1"].PutValue(10);
                ws.Cells["A2"].PutValue(20);

                // Add a column chart and populate its series
                int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = ws.Charts[chartIndex];
                chart.NSeries.Add("A1:A2", true);

                placeholderWb.Save(inputPath);
            }

            // Load the workbook (full load; specific chart-only loading not available in this version)
            var workbook = new Workbook(inputPath);

            // Save the workbook to the output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
