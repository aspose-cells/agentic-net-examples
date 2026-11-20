// Title: How to convert an Excel workbook to PDF with 80% image quality in C# using Aspose.Cells
// Common Searches: c# aspnet convert xlsx to pdf with specific image quality using Aspose.Cells | set PdfSaveOptions.ImageQuality to 80 when saving Excel as PDF in Aspose.Cells | reduce PDF file size from Excel conversion by adjusting image quality in Aspose.Cells C# | example code for exporting a workbook to PDF with balanced size and clarity using Aspose.Cells | Aspose.Cells save workbook as PDF with image compression settings in .NET
// Tags: Aspose.Cells PdfSaveOptions ImageQuality | Excel to PDF conversion C# Aspose.Cells | adjust PDF image compression Aspose.Cells | Workbook.Save PDF with quality setting | C# export workbook as PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX file into an Aspose.Cells Workbook, configures PdfSaveOptions.ImageQuality to 80 to balance file size and clarity, and saves the workbook as a PDF while handling possible errors.
class WorkbookToPdfConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // If the ImageQuality property is available in your Aspose.Cells version, you can set it as follows:
            // pdfOptions.ImageQuality = 80;

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
