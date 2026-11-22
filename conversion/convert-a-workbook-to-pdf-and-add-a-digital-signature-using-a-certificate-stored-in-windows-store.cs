// Title: Convert an Excel workbook to PDF with Aspose.Cells in C# (signing requires Aspose.Pdf and a Windows certificate)
// Common Searches: how to export .xlsx to .pdf using Aspose.Cells in C# | C# example for converting Excel workbook to PDF with Aspose | Aspose.Cells save workbook as PDF then apply digital signature with Aspose.Pdf | load X509Certificate from Windows certificate store for PDF signing in C# | add a digital signature to a PDF generated from Excel using Windows store certificate
// Tags: Aspose.Cells API Excel PDF export C# | Aspose.Pdf X509 digital signing C# | Windows store X509 certificate retrieval C# | Workbook to PDF conversion using Aspose.Cells | PDF signing after Excel export with Aspose.Pdf

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, converts and saves it as a PDF, and explains that adding a digital signature requires the Aspose.Pdf library together with an X509 certificate retrieved from the Windows certificate store.
class WorkbookToPdfWithSignature
{
    static void Main()
    {
        try
        {
            // Define file paths
            const string inputPath = "input.xlsx";
            const string pdfPath = "output.pdf";

            // Verify that the input Excel file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the Excel workbook
            var workbook = new Workbook(inputPath);

            // Convert and save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully converted to PDF: {pdfPath}");

            // NOTE: PDF signing requires Aspose.Pdf library, which is not referenced in this project.
            // If PDF signing is needed, add the Aspose.Pdf NuGet package and implement signing logic here.

            // Cleanup
            workbook.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
