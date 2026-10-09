// Title: C# code to load an XLSB workbook with Aspose.Cells and export it to PDF using default PdfSaveOptions
// AI Prompts: Write a C# program that verifies an XLSB file path, loads the workbook with Aspose.Cells LoadOptions, and saves it as a PDF with basic PdfSaveOptions. | Demonstrate how to configure Aspose.Cells LoadOptions for the XLSB format before performing a PDF conversion in .NET. | Create a C# snippet that catches and logs any exceptions occurring during the XLSB‑to‑PDF conversion with Aspose.Cells.
// Common Searches: asp.net convert xlsb workbook to pdf using aspose.cells c# example | c# loadoptions xlsb format aspose.cells documentation | how to use pdfsaveoptions for simple pdf export in aspose.cells c# | c# check file existence before converting workbook to pdf aspose.cells
// Tags: XLSB workbook loading Aspose.Cells | PDF export PdfSaveOptions Aspose.Cells | C# file existence validation for conversion | exception handling Aspose.Cells workbook save | basic PDF conversion Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The program checks that an XLSB file exists, loads it with Aspose.Cells using LoadOptions set to the XLSB format, and saves it as a PDF with default PdfSaveOptions while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the source XLSB workbook
        string sourcePath = @"C:\Input\Workbook.xlsb";

        // Path for the output PDF file
        string outputPath = @"C:\Output\Workbook.pdf";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the XLSB workbook with appropriate load options
            var loadOptions = new LoadOptions(LoadFormat.Xlsb);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Configure PDF save options (basic options; advanced properties like PreserveOfficeAddIns are not available)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
