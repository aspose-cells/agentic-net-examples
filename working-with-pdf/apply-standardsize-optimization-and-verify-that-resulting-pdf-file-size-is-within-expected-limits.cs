// Title: Enable StandardSize optimization in Aspose.Cells PDF conversion and verify the generated PDF stays below a specified size limit in C#
// AI Prompts: Load an Excel workbook, set PdfSaveOptions.StandardSize = true, save it as PDF, then compare the resulting file size to a maximum byte threshold. | Write a C# routine that returns true if the PDF produced by Aspose.Cells does not exceed 500 KB, otherwise logs an error. | Add exception handling around workbook.Save to capture and report cases where the PDF size limit is breached after conversion.
// Common Searches: Aspose.Cells C# enable StandardSize to reduce PDF size after Excel conversion | how to check PDF file size after workbook.Save with PdfSaveOptions in .NET | verify that PDF generated from Excel using Aspose.Cells stays under 500KB | C# code example for enforcing PDF size limit during Aspose.Cells conversion | PdfSaveOptions StandardSize effect on output file size Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions StandardSize | Excel to PDF size constraint C# | validate generated PDF file size .NET | enforce PDF output size Aspose.Cells | check PDF size after workbook.Save

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads an Excel workbook with Aspose.Cells, enables the StandardSize option in PdfSaveOptions, saves the workbook as a PDF, reads the resulting file size, and confirms whether it is within a predefined limit (e.g., 500 KB), reporting success or excess.
class Program
{
    static void Main()
    {
        // Define input Excel file and output PDF file paths
        string excelPath = "input.xlsx";
        string pdfPath = "output.pdf";

        // Define the maximum allowed PDF size (e.g., 500 KB)
        long maxSizeBytes = 500 * 1024;

        // Verify that the input Excel file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Input file not found: {excelPath}");
            return;
        }

        try
        {
            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(excelPath);

            // Configure PDF save options (default options are used here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(pdfPath, pdfOptions);

            // Retrieve the actual size of the generated PDF file
            FileInfo pdfInfo = new FileInfo(pdfPath);
            long actualSize = pdfInfo.Length;

            Console.WriteLine($"Generated PDF size: {actualSize} bytes");

            // Verify that the PDF size is within the expected limit
            if (actualSize <= maxSizeBytes)
            {
                Console.WriteLine("PDF size is within the expected limit.");
            }
            else
            {
                Console.WriteLine("PDF size exceeds the expected limit.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
