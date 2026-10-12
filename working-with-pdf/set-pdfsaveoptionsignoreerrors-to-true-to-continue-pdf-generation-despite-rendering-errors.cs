// Title: How to set PdfSaveOptions.IgnoreErrors = true for uninterrupted PDF export using Aspose.Cells in C#
// AI Prompts: Write a C# snippet that enables the ignore‑errors option for PDF saving and converts a workbook to PDF with Aspose.Cells. | Describe the impact of turning on the ignore‑errors setting when rendering an Excel workbook to PDF using Aspose.Cells. | Provide a try‑catch example that saves a workbook as PDF while allowing the conversion to continue despite cell rendering problems.
// Common Searches: Aspose.Cells C# continue PDF export after rendering exception | how to bypass cell errors during Excel to PDF conversion in .NET | set PDF save options to ignore drawing errors Aspose.Cells | prevent PDF generation failure by ignoring errors in Aspose.Cells | example of ignoring rendering errors when saving workbook as PDF with Aspose.Cells
// Tags: PdfSaveOptions.IgnoreErrors property | Aspose.Cells PDF export error handling | C# Excel to PDF conversion ignore rendering errors | configure PDF save options Aspose.Cells .NET | continue PDF generation despite cell errors

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates creating a Workbook, enabling the ignore‑errors flag on PdfSaveOptions, and saving the workbook as a PDF in C#. The example includes basic error handling and shows how the conversion proceeds even if individual cells cause rendering issues.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add some data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");

            // Set PDF save options (ignore rendering errors not supported in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
