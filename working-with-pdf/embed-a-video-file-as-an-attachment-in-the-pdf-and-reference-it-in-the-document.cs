// Title: Create a PDF from an Excel workbook with a clickable link to a local MP4 video using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new Workbook, places a prompt in cell A1, adds a hyperlink to a local MP4 file, and saves the workbook as a PDF with Aspose.Cells so the video can be opened from the PDF. | Demonstrate how to verify that a video file exists before inserting a hyperlink into an Excel worksheet and then exporting the sheet to PDF with Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells to generate a PDF that contains a reference to an embedded video file.
// Common Searches: asp.net add hyperlink to local mp4 in Excel sheet and export as PDF with Aspose.Cells | c# embed video link in PDF generated from Excel using Aspose.Cells | check file existence before creating hyperlink in Aspose.Cells workbook before PDF conversion
// Tags: add hyperlink to mp4 in Excel Aspose.Cells | export workbook to PDF with media reference Aspose.Cells | file existence validation before PDF export C# | PdfSaveOptions media link configuration Aspose.Cells | embed video reference in PDF from Excel

using System;
using System.IO;
using Aspose.Cells;

// The example checks for a local MP4 file, creates a new workbook, writes a prompt in cell A1, adds a hyperlink to the video, and saves the workbook as a PDF using Aspose.Cells' PdfSaveOptions, producing a PDF that references the video file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the video file
            string videoPath = "sample.mp4";

            // Ensure the video file exists before proceeding
            if (!File.Exists(videoPath))
            {
                Console.WriteLine($"Video file not found: {videoPath}");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Write a prompt in cell A1
            sheet.Cells["A1"].PutValue("Click the link to open the video");

            // Add a hyperlink in cell A1 that points to the video file (single cell range)
            sheet.Hyperlinks.Add(0, 0, 1, 1, videoPath);

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            string outputPdf = "DocumentWithVideo.pdf";
            workbook.Save(outputPdf, pdfOptions);
            Console.WriteLine($"PDF saved successfully: {outputPdf}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
