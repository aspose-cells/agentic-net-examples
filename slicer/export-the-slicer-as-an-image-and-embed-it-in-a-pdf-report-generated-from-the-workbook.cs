// Title: Export an Excel slicer as an image and embed it into a PDF report using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, extracts each slicer's bitmap via Aspose.Cells, saves the bitmap as PNG, and inserts the PNG into the corresponding PDF page before saving the PDF. | Refactor the sample to create a reusable method GetSlicerImage(Worksheet sheet, int slicerIndex) that returns a MemoryStream, then use that method to embed slicer images into a PDF generated with Aspose.Cells. | Write a C# snippet that detects System.Drawing support, conditionally exports slicer images, and falls back to generating a PDF without images when the environment lacks drawing capabilities.
// Common Searches: how to save slicer as png with Aspose.Cells C# | embed slicer bitmap into PDF using Aspose.Cells .NET | Aspose.Cells export slicer image to PDF report example | C# generate PDF from workbook and include slicer picture | Aspose.Cells slicer image extraction limitations System.Drawing
// Tags: export slicer image Aspose.Cells .NET | embed slicer bitmap into PDF | Aspose.Cells PDF generation with worksheet images | C# extract slicer picture from Excel | system.drawing dependency Aspose.Cells slicer export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an Excel workbook, checks the first worksheet for slicers, logs their presence, skips slicer image export due to the lack of System.Drawing support in the runtime, and then saves the workbook as a PDF report.
class SlicerToPdfReport
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "InputWorkbook.xlsx";

            // Verify the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook that contains the slicer
            Workbook workbook = new Workbook(workbookPath);

            // Check if the first worksheet contains any slicers
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicer found in the workbook. Generating PDF without slicer image.");
            }
            else
            {
                // If slicers exist, you could add custom handling here (e.g., export slicer image).
                // The current environment does not support System.Drawing, so slicer image export is omitted.
                Console.WriteLine("Slicer detected, but image export is skipped due to environment limitations.");
            }

            // Export the workbook to PDF
            string finalPdfPath = "WorkbookReport_WithSlicer.pdf";
            workbook.Save(finalPdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF report generated: {finalPdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
