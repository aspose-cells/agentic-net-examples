// Title: Apply StandardSize PDF Optimization in Aspose.Cells (.NET) and Validate File Size
// Description: C# sample that creates a workbook, saves it as a PDF using PdfSaveOptions with OptimizationType.Standard (high‑print quality), then checks that the generated file stays under a 1 MB limit.
// Keywords: Aspose.Cells PDF optimization | PdfSaveOptions Standard | StandardSize PDF | C# PDF size verification | PdfOptimizationType.Standard | reduce PDF file size Aspose | high quality printable PDF .NET | file size limit check | Aspose.Cells sample code | PDF generation with size constraint
// Common Searches: Aspose.Cells StandardSize PDF optimization example C# | How to verify PDF file size after saving with Aspose.Cells | PdfSaveOptions OptimizationType.Standard usage | Check PDF size limit in .NET after Aspose.Cells export | C# code to enforce PDF size under 1 MB with Aspose
// Developer Intent: Save a workbook as a PDF using StandardSize optimization and ensure the output does not exceed a predefined size threshold.
// Use Cases: Produce high‑quality printable PDFs for reports while keeping each file under 1 MB. | Automate batch conversion of workbooks to PDF with size validation to flag oversized documents. | Integrate PDF size checks into CI/CD pipelines to prevent large PDFs from reaching production.
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook to PDF with StandardSize optimization and throws an exception if the file exceeds 2 MB. | Compare PdfOptimizationType.Standard, MinimumSize, and MaximumQuality, and suggest the best scenario for each. | Provide additional techniques (e.g., image compression, font subsetting) to further shrink a PDF after applying StandardSize optimization with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# sample that creates a workbook, saves it as a PDF using PdfSaveOptions with OptimizationType.Standard (high‑print quality), then checks that the generated file stays under a 1 MB limit.
class PdfStandardSizeOptimizationDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        for (int i = 2; i <= 101; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Product {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Configure PDF save options to use Standard optimization (high print quality)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.Standard; // StandardSize optimization

        // Define output PDF path
        string pdfPath = "StandardOptimizedOutput.pdf";

        // Save the workbook as PDF with the specified options
        workbook.Save(pdfPath, pdfOptions);

        // Verify the resulting PDF file size
        FileInfo pdfInfo = new FileInfo(pdfPath);
        long fileSizeBytes = pdfInfo.Length;
        // Example expected limit: 1 MB (1,048,576 bytes)
        long expectedLimitBytes = 1_048_576;

        Console.WriteLine($"PDF saved to: {pdfPath}");
        Console.WriteLine($"File size: {fileSizeBytes} bytes");

        if (fileSizeBytes <= expectedLimitBytes)
        {
            Console.WriteLine("File size is within the expected limit.");
        }
        else
        {
            Console.WriteLine("File size exceeds the expected limit.");
        }
    }
}
