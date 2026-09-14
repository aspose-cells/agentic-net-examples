// Title: Convert an XLSX workbook to PDF with Aspose.Cells for .NET and validate the generated file size
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells and saves it as a PDF with the default SaveFormat.Pdf option. | Add logic to the conversion program that checks whether the resulting PDF file exists and prints its size in bytes. | Modify the sample to accept source and destination file paths as command‑line arguments while preserving the file‑size verification step.
// Common Searches: Aspose.Cells C# example for converting Excel to PDF with default settings | How to check PDF file size after saving workbook using Aspose.Cells .NET | Command line tool to convert XLSX to PDF using Aspose.Cells and report output size | Validate that PDF generated from Excel is not empty with Aspose.Cells in C#
// Tags: Aspose.Cells XLSX to PDF conversion | C# verify generated PDF size | default PDF save options Aspose.Cells | command‑line file path parameters Aspose.Cells | PDF output existence check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads an XLSX workbook with Aspose.Cells, saves it as a PDF using default settings, then confirms the PDF file exists and outputs its byte size.
class XlsxToPdfConverter
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Path for the generated PDF file
        string outputPath = "output.pdf";

        // Load the XLSX workbook (default load options)
        Workbook workbook = new Workbook(inputPath);

        // Save the workbook as PDF using default save options
        workbook.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created and check its size
        FileInfo pdfInfo = new FileInfo(outputPath);
        if (pdfInfo.Exists && pdfInfo.Length > 0)
        {
            Console.WriteLine($"PDF successfully created. File size: {pdfInfo.Length} bytes.");
        }
        else
        {
            Console.WriteLine("PDF creation failed or file is empty.");
        }
    }
}
