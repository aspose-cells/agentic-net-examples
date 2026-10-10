// Title: Convert an XLSX workbook to PDF with Aspose.Cells for .NET and check the generated file size
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, saves it as a PDF with the default SaveFormat.Pdf, and prints the PDF file size in bytes. | Show how to obtain the size of a PDF file created from an Excel workbook after conversion using .NET file I/O.
// Common Searches: asp.net core convert xlsx to pdf with aspose.cells and get file size | c# Aspose.Cells save workbook as PDF default options example | how to read PDF file size after converting Excel to PDF in C# | Aspose.Cells PDF conversion verify output size programmatically | sample code to export Excel to PDF and check file length using Aspose.Cells
// Tags: Aspose.Cells XLSX to PDF conversion | default PDF save options Aspose.Cells | retrieve generated PDF file size .NET | C# workbook export to PDF example | validate PDF output size programmatically

using System;
using System.IO;
using Aspose.Cells;

// The program loads an XLSX workbook with Aspose.Cells, saves it as a PDF using the default SaveFormat.Pdf, then reads and displays the resulting PDF file's size in bytes.
class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string inputPath = "input.xlsx";

        // Path where the PDF will be saved
        string outputPath = "output.pdf";

        // Load the XLSX workbook
        Workbook workbook = new Workbook(inputPath);

        // Convert and save the workbook to PDF using default settings
        workbook.Save(outputPath, SaveFormat.Pdf);

        // Verify the size of the generated PDF file
        FileInfo pdfInfo = new FileInfo(outputPath);
        long pdfSize = pdfInfo.Length;

        Console.WriteLine($"PDF file created at: {outputPath}");
        Console.WriteLine($"PDF file size: {pdfSize} bytes");
    }
}
