// Title: Convert an XLSX workbook to PDF with Aspose.Cells for .NET and check the generated file size
// Common Searches: C# convert Excel .xlsx to PDF using Aspose.Cells and get file size | How to save an Aspose.Cells workbook as PDF with default options in .NET | Retrieve PDF output size after converting workbook with Aspose.Cells | Aspose.Cells .NET example for XLSX to PDF conversion and size verification | Check generated PDF file length after Aspose.Cells SaveFormat.Pdf
// Tags: Aspose.Cells XLSX to PDF conversion | Workbook.Save with SaveFormat.Pdf in C# | Verify PDF file size after conversion | Default PDF settings Aspose.Cells .NET | Measure output PDF size using FileInfo

using System;
using System.IO;
using Aspose.Cells;

// The program loads an XLSX workbook, saves it as a PDF using Aspose.Cells' default PDF settings, then reads the resulting PDF file's size in bytes and prints the file path and size to the console.
class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Path for the generated PDF file
        string pdfPath = "output.pdf";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as PDF using default settings
        workbook.Save(pdfPath, SaveFormat.Pdf);

        // Verify the output file size
        FileInfo pdfInfo = new FileInfo(pdfPath);
        long fileSizeBytes = pdfInfo.Length;

        Console.WriteLine($"PDF file created at: {pdfPath}");
        Console.WriteLine($"File size: {fileSizeBytes} bytes");
    }
}
