// Title: Embedding an image into a PDF generated from an Excel workbook using Aspose.Cells PdfSaveOptions in C#
// Common Searches: C# Aspose.Cells insert image into worksheet and save as PDF | How to embed a PNG file in PDF output when converting Excel with Aspose.Cells | Using PdfSaveOptions to include images in generated PDF from workbook | Save Excel workbook as PDF with embedded picture using Aspose.Cells C# | Aspose.Cells convert Excel to PDF with attached image example
// Tags: picture insertion before PDF export Aspose.Cells | PdfSaveOptions image embedding C# | Excel worksheet image to PDF conversion | Aspose.Cells embed PNG in PDF | generate PDF with workbook image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds data, inserts a PNG image into the worksheet, and then saves the workbook as a PDF using PdfSaveOptions, resulting in a PDF that contains the embedded image.
class EmbedImageAsPdfAttachment
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Optionally add some data to the workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);

            // Path to the image file that will be embedded in the PDF
            string imagePath = @"C:\Images\sampleImage.png";

            // Ensure the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine("Image file not found: " + imagePath);
                return;
            }

            // Insert the image into the worksheet (it will appear in the PDF)
            // Parameters: upper-left row, upper-left column, image file path
            sheet.Pictures.Add(2, 0, imagePath);

            // Ensure the output directory exists
            string outputPdfPath = @"C:\Output\WorkbookWithImage.pdf";
            string outputDir = Path.GetDirectoryName(outputPdfPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF (image will be part of the PDF content)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save(outputPdfPath, pdfOptions);

            Console.WriteLine("PDF saved with embedded image at: " + outputPdfPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
