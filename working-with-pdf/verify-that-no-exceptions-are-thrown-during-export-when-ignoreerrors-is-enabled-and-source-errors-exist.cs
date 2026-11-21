// Title: Export a workbook with #REF! formula errors to PDF without throwing exceptions using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export to PDF while ignoring formula errors in C# | C# save Excel workbook as PDF when cells contain #REF! errors without exception | How to enable IgnoreErrors in Aspose.Cells to prevent PDF conversion failures | Prevent exceptions during PDF conversion of Excel files with invalid references using Aspose.Cells | Aspose.Cells .NET PDF generation with cells that return error values
// Tags: ignore errors PDF export Aspose.Cells | export workbook with formula errors to PDF | Aspose.Cells SaveFormat.Pdf error handling | C# Aspose.Cells ignore formula errors | PDF conversion of Excel #REF! cells

using System;
using Aspose.Cells;

// The example creates a workbook, inserts a formula that produces a #REF! error, and saves the workbook as a PDF using Aspose.Cells. The operation is wrapped in a try‑catch block to confirm that no exception is thrown during the PDF export despite the presence of formula errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a formula that will generate an error (#REF!)
            sheet.Cells["A1"].Formula = "=SUM(UNKNOWN!A1:A5)";

            // Export to PDF (Aspose.Cells will render the error value without throwing)
            workbook.Save("ExportedOutput.pdf", SaveFormat.Pdf);

            Console.WriteLine("Export completed successfully without exceptions.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Export failed with exception: " + ex.Message);
        }
    }
}
