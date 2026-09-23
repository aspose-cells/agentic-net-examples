// Title: Save each chart from an Excel workbook as a JPEG file named after its worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, iterates over all worksheets, extracts every chart, and writes each chart to a JPEG file whose name combines the worksheet name and the chart index. | Enhance a C# Aspose.Cells routine that exports charts to JPEG by adding checks for missing workbook files, empty chart collections, and file‑system write permissions, and log meaningful error messages. | Implement a reusable method `ExportCharts(string workbookPath, string outputFolder)` that uses Aspose.Cells to save every chart in the specified workbook as a JPEG image, naming the files `<WorksheetName>_Chart<index>.jpg`.
// Common Searches: aspnet export excel chart to jpeg with worksheet name in filename | c# Aspose.Cells save each chart as separate image file | how to loop through worksheets and charts in Aspose.Cells and export to jpg | generate jpeg images for all charts in an Excel file using Aspose.Cells .NET | save chart images with custom naming convention Aspose.Cells C#
// Tags: aspocells chart image extraction | c# iterate workbook charts | jpeg output naming convention aspocells | imageorprintoptions chart rendering | batch chart image generation aspocells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// // Loads 'input.xlsx', iterates through every worksheet and its charts, and saves each chart as a JPEG file named '<WorksheetName>_Chart<index>.jpg' using Aspose.Cells.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Define image options for exporting the chart
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                        {
                            // OnePagePerSheet is not used for charts but kept for consistency
                            OnePagePerSheet = true
                        };

                        // Determine a unique file name for the chart image
                        int chartIndex = sheet.Charts.IndexOf(chart);
                        string outputFileName = $"{sheet.Name}_Chart{chartIndex}.jpg";

                        // Export the chart to an image file (file name first, then options)
                        chart.ToImage(outputFileName, imgOptions);
                        Console.WriteLine($"Chart saved to '{outputFileName}'.");
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Failed to export chart on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
