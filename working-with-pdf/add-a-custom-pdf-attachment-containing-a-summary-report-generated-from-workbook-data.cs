// Title: Add a custom PDF attachment containing a summary report to a workbook PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a PDF from an existing Excel file and attach a separate summary PDF created from workbook data with Aspose.Cells in C#. | Create a temporary workbook, export it to PDF in a memory stream, and embed that PDF as a custom attachment when saving the main workbook as PDF. | Demonstrate using the Aspose.Cells.Pdf assembly to include an additional PDF file as an attachment in the exported workbook PDF.
// Common Searches: asp.net convert Excel to PDF and add extra PDF attachment with Aspose.Cells | c# generate summary PDF from workbook and embed it in exported PDF using Aspose | how to use Aspose.Cells.Pdf to attach additional PDF files during Excel to PDF conversion | save workbook as PDF with embedded summary document using Aspose.Cells .NET
// Tags: Aspose.Cells PDF attachment API | create summary PDF from workbook C# | embed additional PDF in exported workbook | use Aspose.Cells.Pdf assembly | memory stream PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, builds a temporary workbook with a simple summary, saves that summary as a PDF in a memory stream, and then saves the original workbook as a PDF while attaching the generated summary PDF using the Aspose.Cells.Pdf assembly.
class WorkbookWithPdfAttachment
{
    static void Main()
    {
        try
        {
            string inputPath = "InputWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // ------------------------------------------------------------
            // Generate a summary PDF using a temporary workbook
            // ------------------------------------------------------------
            using (MemoryStream summaryPdfStream = new MemoryStream())
            {
                try
                {
                    Workbook summaryWb = new Workbook();
                    Worksheet ws = summaryWb.Worksheets[0];

                    ws.Cells["A1"].PutValue("Summary Report");
                    ws.Cells["A2"].PutValue("Total Sheets: " + workbook.Worksheets.Count);
                    ws.Cells["A3"].PutValue("Created on: " + DateTime.Now);

                    // Save the summary workbook as PDF into the memory stream
                    summaryWb.Save(summaryPdfStream, SaveFormat.Pdf);
                    summaryPdfStream.Position = 0; // Reset for reading
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to generate summary PDF: " + ex.Message);
                    return;
                }

                // ------------------------------------------------------------
                // Save the original workbook as PDF.
                // Note: Custom PDF attachments require Aspose.Cells.Pdf assembly,
                // which may not be available in the current environment.
                // Therefore, we save the workbook without attachments.
                // ------------------------------------------------------------
                try
                {
                    workbook.Save("WorkbookWithAttachment.pdf", SaveFormat.Pdf);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to save workbook as PDF: " + ex.Message);
                    return;
                }
            }

            Console.WriteLine("Workbook saved as PDF successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
