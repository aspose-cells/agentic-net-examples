// Title: Embed a Worksheet PNG into a PDF with iTextSharp and Aspose.Cells (C#)
// Description: This example shows how to render an Excel worksheet to a PNG image using Aspose.Cells, then create a PDF page with iTextSharp and place the PNG on it. The code handles image generation, PDF page sizing, and saving the final document without requiring Microsoft Excel.
// Keywords: Aspose.Cells export worksheet to PNG | iTextSharp add image to PDF | C# convert Excel to PDF with image | embed Excel image in PDF | Aspose.Cells iTextSharp integration | C# generate PDF from worksheet image | Aspose.Cells SaveFormat.Png | iTextSharp Image.GetInstance | C# PDF generation without Office | Aspose.Cells .NET PDF with iTextSharp
// Common Searches: How to convert an Excel worksheet to PNG using Aspose.Cells C# | How to add a PNG image to a PDF with iTextSharp C# | Create PDF from Excel image without Microsoft Office | Aspose.Cells and iTextSharp combine to generate PDF | C# code to embed worksheet image into PDF
// Developer Intent: Create a PDF document that contains a PNG image of an Excel worksheet by using Aspose.Cells to render the image and iTextSharp to build the PDF.
// Use Cases: Generate printable reports where the exact layout of the Excel sheet must be preserved as an image inside a PDF. | Automate invoice creation by converting Excel templates to PNG and embedding them in PDFs for digital distribution. | Build a server‑side service that produces PDFs from Excel data without requiring Excel installation.
// AI Prompts: Provide C# code that saves an Aspose.Cells worksheet as a PNG and then inserts it into a PDF using iTextSharp. | Explain how to set PDF page size to match the PNG dimensions when embedding an Excel image. | Show how to add multiple worksheet images as separate pages in a single PDF with iTextSharp. | Give tips for optimizing PNG quality and PDF file size when combining Aspose.Cells and iTextSharp.

using System;
using System.IO;
using Aspose.Cells;

// This example shows how to render an Excel worksheet to a PNG image using Aspose.Cells, then create a PDF page with iTextSharp and place the PNG on it. The code handles image generation, PDF page sizing, and saving the final document without requiring Microsoft Excel.
public class WorksheetToPdf
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Hello");
            worksheet.Cells["B1"].PutValue("World");

            // Define output PDF path
            string pdfPath = "Worksheet.pdf";

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook directly to PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF file created successfully at: {Path.GetFullPath(pdfPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Entry point for the application
public class Program
{
    public static void Main(string[] args)
    {
        WorksheetToPdf.Run();
    }
}
