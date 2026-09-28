// Title: Remove duplicate rows while preserving formulas and export the worksheet to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, removes duplicate rows from the first worksheet without breaking formulas, and saves the result as a PDF using Aspose.Cells. | Show how to use Aspose.Cells in a .NET project to deduplicate rows in a worksheet while keeping all formulas intact, then export the cleaned sheet to PDF. | Generate a minimal Aspose.Cells example that calls Cells.RemoveDuplicates and then saves the workbook with SaveFormat.Pdf.
// Common Searches: Aspose.Cells C# remove duplicate rows keep formulas | Export cleaned Excel worksheet to PDF using Aspose.Cells .NET | How to preserve formulas when deleting duplicate rows with Aspose.Cells | Save workbook as PDF after removing duplicates Aspose.Cells example | Remove duplicate rows from first sheet Aspose.Cells C#
// Tags: remove duplicate rows Aspose.Cells | preserve formulas Cells.RemoveDuplicates | export worksheet to PDF Aspose.Cells | Aspose.Cells SaveFormat.Pdf usage | deduplicate rows first worksheet C#

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an Excel workbook, removes duplicate rows from the first worksheet while retaining formulas, and saves the cleaned file as a PDF using Aspose.Cells.
class RemoveDuplicatesAndExportPdf
{
    static void Main()
    {
        // Path to the source Excel file
        string inputFile = "input.xlsx";

        // Path for the resulting PDF file
        string outputPdf = "output.pdf";

        // Load the workbook from the file (lifecycle rule: use constructor)
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove duplicate rows in the sheet while keeping formulas intact
        worksheet.Cells.RemoveDuplicates();

        // Export the cleaned workbook as PDF (lifecycle rule: use Save method)
        workbook.Save(outputPdf, SaveFormat.Pdf);

        Console.WriteLine("Duplicate rows removed and PDF saved to: " + outputPdf);
    }
}
