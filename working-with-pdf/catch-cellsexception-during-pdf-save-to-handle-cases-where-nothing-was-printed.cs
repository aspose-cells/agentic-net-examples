// Title: How to Catch CellsException When Saving a Workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a Workbook as PDF inside a try‑catch block and logs the CellsException message when the workbook has no printable content. | Show how to differentiate a CellsException caused by an empty workbook from other exceptions during PDF export using Aspose.Cells. | Provide a pattern for wrapping Aspose.Cells PDF conversion in a reusable method that returns a success flag and captures CellsException details.
// Common Searches: Aspose.Cells how to handle CellsException when exporting to PDF in C# | C# catch empty workbook error during PDF conversion with Aspose.Cells | PDF save fails with no printable area Aspose.Cells .NET exception handling | try‑catch pattern for Aspose.Cells PDF export errors | detect and log CellsException for PDF generation in Aspose.Cells applications
// Tags: CellsException handling for PDF export Aspose.Cells | C# try‑catch Aspose.Cells PDF conversion | empty workbook detection before PDF save Aspose.Cells | Aspose.Cells PDF generation error handling .NET | logging CellsException message during PDF save

using System;
using Aspose.Cells;

// The example creates a Workbook, optionally adds data, attempts to save it as a PDF, and catches CellsException to manage situations where the PDF generation fails—such as when the workbook contains no printable content—while outputting the exception details for further handling.
class Program
{
    static void Main()
    {
        // Create a new workbook (could also be loaded from a file)
        Workbook workbook = new Workbook();

        // Example: add a worksheet with some data (optional)
        // Worksheet sheet = workbook.Worksheets[0];
        // sheet.Cells["A1"].PutValue("Sample");

        try
        {
            // Attempt to save the workbook as PDF
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
        catch (CellsException ex)
        {
            // Handle cases where nothing was printed (e.g., empty workbook)
            Console.WriteLine("PDF save failed: " + ex.Message);
            // Additional error handling can be placed here
        }
    }
}
