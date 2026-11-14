// Title: Generate a PDF containing only a slicer region by setting the worksheet print area with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet print area to slicer range before exporting to PDF | How to export only the slicer area of an Excel sheet to PDF using Aspose.Cells .NET | C# code to define print area for a slicer and save the worksheet as PDF with Aspose.Cells | Generate PDF report of a specific slicer range in Excel using Aspose.Cells for .NET | Configure page setup orientation and fit-to-page for slicer PDF export in Aspose.Cells C#
// Tags: set worksheet print area Aspose.Cells C# | export slicer range to PDF Aspose.Cells | configure page setup for PDF export Aspose.Cells | save workbook as PDF specific range Aspose.Cells | Aspose.Cells slicer PDF generation C#

using System;
using System.IO;
using Aspose.Cells;

namespace SlicerPdfReport
{
    // The example loads an existing Excel workbook, selects the first worksheet, defines the cell range that contains the slicer (B2:E10), sets this range as the worksheet's print area, adjusts orientation and scaling, and then saves the worksheet as a PDF. The resulting PDF includes only the slicer region.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the input workbook
                string inputPath = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or use a specific name)
                Worksheet sheet = workbook.Worksheets[0];

                // Define the slicer region (adjust as needed)
                string slicerRange = "B2:E10";

                // Set the print area to the slicer region only
                sheet.PageSetup.PrintArea = slicerRange;

                // Ensure the page orientation and scaling are appropriate
                sheet.PageSetup.Orientation = PageOrientationType.Portrait;
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 1;

                // Path for the output PDF
                string outputPath = "SlicerReport.pdf";

                // Save the worksheet as a PDF; only the defined print area will be exported
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"PDF report generated successfully at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
