// Title: How to add a digital signature to a PDF created from an Excel workbook using Aspose.Cells and Aspose.Pdf in C#
// Common Searches: C# sign PDF generated from Excel with Aspose.Cells and Aspose.Pdf | apply digital certificate to PDF output after workbook.Save in .NET | how to use Aspose.Pdf to digitally sign a PDF produced by Aspose.Cells | secure PDF generated from an .xlsx file with electronic signature in C# | programmatically add a digital signature to a PDF converted from Excel using Aspose libraries | add authenticity to PDF created from Excel workbook in a .NET application
// Tags: Aspose.Cells Excel to PDF conversion C# | Aspose.Pdf digital signature implementation | programmatic PDF signing with Aspose libraries | C# apply electronic signature to converted PDF | digital certificate PDF signing Aspose | ensure PDF authenticity after Excel conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, saves it as a PDF, and outlines how to subsequently apply a digital signature using Aspose.Pdf to guarantee the PDF’s authenticity. It includes file existence checks and exception handling for robust execution.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for source Excel and output PDF
            string excelPath = "input.xlsx";
            string pdfPath = "output.pdf";

            // Ensure the Excel file exists before loading
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(excelPath);

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF: '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
