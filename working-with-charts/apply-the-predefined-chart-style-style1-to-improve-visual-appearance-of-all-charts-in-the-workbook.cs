// Title: Apply the built‑in Style1 chart style to every chart in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, iterates through all worksheets and charts, sets each chart's Style property to 1 (Style1), and saves the workbook. | Show how to bulk‑apply a predefined chart style to all charts in a workbook using the Aspose.Cells Chart.Style API. | Create a script that checks for the input file, creates missing output directories, and updates every chart to Style1 before writing the result.
// Common Searches: asp.net aspose.cells apply built‑in chart style to all charts in workbook | c# set chart.Style = 1 for multiple charts using Aspose.Cells | how to programmatically change Excel chart appearance to Style1 with Aspose.Cells
// Tags: apply predefined chart style Aspose.Cells C# | set chart.Style property for all charts .NET | bulk update Excel chart appearance using Aspose.Cells | iterate worksheets and charts Aspose.Cells API | use chart style index 1 in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads input.xlsx with Aspose.Cells, loops through each worksheet and its charts, assigns chart.Style = 1 (the built‑in Style1), ensures the output folder exists, and saves the modified workbook as output.xlsx while handling missing files and exceptions.
class ApplyChartStyle
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and apply chart style
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Apply a predefined chart style (Style1 corresponds to index 1)
                    chart.Style = 1;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
