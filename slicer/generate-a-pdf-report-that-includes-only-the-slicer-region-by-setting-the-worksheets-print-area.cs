// Title: Export only the slicer region to PDF by configuring the worksheet print area with Aspose.Cells for .NET (C#)
// AI Prompts: Assign the slicer's cell range to Worksheet.PageSetup.PrintArea and save the workbook as a PDF using Aspose.Cells in C#. | Programmatically retrieve the slicer's address, set it as the print area, and generate a single‑page PDF report. | Configure FitToPagesWide and FitToPagesTall so the slicer area fits on one PDF page during export.
// Common Searches: Aspose.Cells C# export slicer only to PDF | set print area to slicer range before saving as PDF using Aspose.Cells | C# generate PDF report that includes just the slicer region in an Excel workbook | Aspose.Cells page scaling for slicer export to PDF
// Tags: Aspose.Cells set worksheet print area C# | export slicer range to PDF Aspose.Cells | Aspose.Cells PDF page scaling slicer | C# define print area for slicer | Aspose.Cells save workbook as PDF limited range

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, defines the slicer’s cell range (e.g., A1:D20) as the worksheet PrintArea, sets FitToPagesWide and FitToPagesTall to 1 to fit the area on a single page, and saves the result as a PDF using Aspose.Cells for .NET, with basic file existence checks and exception handling.
class PdfReportWithSlicer
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string inputFile = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the slicer region (example: cells A1 to D20)
            // Replace with the actual slicer range in your worksheet
            string slicerRange = "A1:D20";

            // Set the print area to the slicer region
            sheet.PageSetup.PrintArea = slicerRange;

            // Optionally, set the page scaling (orientation line removed due to API differences)
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 1;

            // Save the workbook as a PDF; only the defined print area will be included
            string outputPdf = "SlicerReport.pdf";
            workbook.Save(outputPdf, SaveFormat.Pdf);

            Console.WriteLine($"PDF report generated successfully: {outputPdf}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
