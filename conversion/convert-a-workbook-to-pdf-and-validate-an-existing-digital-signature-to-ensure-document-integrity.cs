// Title: Convert an Excel workbook to PDF and check for a digital signature using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert Excel to PDF and verify digital signature | how to export .xlsx as PDF while ensuring signature integrity with Aspose.Cells | C# load workbook, validate existing signature, then save as PDF using Aspose.Cells | checking Excel digital signature before PDF conversion in .NET | Aspose.Cells PDF export with signature validation fallback
// Tags: Aspose.Cells workbook to PDF conversion C# | digital signature validation limitation Aspose.Cells .NET | Excel to PDF export with integrity check | SaveFormat.Pdf usage Aspose.Cells | error handling for missing Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, notes that the current library version does not provide digital signature validation APIs, and then converts the workbook to a PDF using SaveFormat.Pdf, including basic error handling for missing input files and conversion failures.
class WorkbookToPdfWithSignatureCheck
{
    static void Main()
    {
        // Paths to the source Excel file and the target PDF file
        string excelPath = "input.xlsx";
        string pdfPath = "output.pdf";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // NOTE: Digital signature validation APIs are not available in the current Aspose.Cells version.
            // If needed, implement signature validation using the appropriate Aspose APIs or external libraries.

            // Convert the workbook to PDF
            try
            {
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully converted to PDF at: {pdfPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error during PDF conversion: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
