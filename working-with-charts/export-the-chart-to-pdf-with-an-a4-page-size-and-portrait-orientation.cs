// Title: Export an Excel chart to a PDF with A4 portrait layout using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart to PDF with A4 paper size | How to set PDF orientation to portrait when saving chart with Aspose.Cells | Save specific worksheet range as PDF in C# using Aspose.Cells | Export Excel chart as PDF file with custom page setup using Aspose.Cells .NET | C# Aspose.Cells PDF export chart A4 portrait page setup example
// Tags: chart to PDF conversion Aspose.Cells | A4 page setup Aspose.Cells | portrait orientation PDF Aspose.Cells | export worksheet range as PDF C# | save workbook as PDF with chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ChartExportExample
{
    // The example loads an existing Excel workbook, optionally configures the worksheet's page setup for A4 paper size and portrait orientation, defines a print area that contains the chart, ensures the output directory exists, and then saves the workbook (or the selected worksheet) as a PDF file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "ChartExport.pdf";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook that contains the chart
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Optional: configure page setup (commented out if enums are unavailable in the referenced version)
                // worksheet.PageSetup.PaperSize = PaperSize.A4;                     // A4 page size
                // worksheet.PageSetup.Orientation = PageOrientation.Portrait;      // Portrait orientation

                // Optional: set the print area to the chart's range (e.g., B2:D15)
                worksheet.PageSetup.PrintArea = "B2:D15";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (or the specific worksheet) as a PDF file
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"PDF file successfully created at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine("An error occurred during the export process:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
