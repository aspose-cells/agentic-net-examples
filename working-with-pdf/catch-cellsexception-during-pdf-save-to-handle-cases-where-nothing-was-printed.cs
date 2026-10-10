// Title: Catch Aspose.Cells CellsException when saving a workbook to PDF to handle 'nothing was printed' errors in C#
// AI Prompts: Generate C# code that saves an Excel workbook as PDF with Aspose.Cells inside a try‑catch block and checks the CellsException message for the phrase "nothing was printed". | Provide a C# example that logs a custom warning and optionally creates a placeholder PDF page when Aspose.Cells throws a CellsException indicating no printable content.
// Common Searches: asp.net catch CellsException during Excel to PDF conversion when workbook has no printable area | how to detect 'nothing was printed' error in Aspose.Cells PDF export C# | handle empty worksheet PDF save exception Aspose.Cells | C# Aspose.Cells save as PDF returns CellsException with no content message
// Tags: catch CellsException PDF export Aspose.Cells | handle empty printable area exception | Aspose.Cells PDF save error handling | C# exception handling for no printable content | detect missing printable content Aspose.Cells

using System;
using Aspose.Cells;

// The program loads an Excel workbook, attempts to save it as a PDF using Aspose.Cells, and catches a CellsException to specifically handle cases where the exception message indicates that nothing was printed (e.g., empty workbook or no printable area).
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define the output PDF file path
        string pdfPath = "output.pdf";

        try
        {
            // Attempt to save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine("PDF saved successfully.");
        }
        catch (CellsException ex)
        {
            // Handle the case where nothing was printed (e.g., empty workbook or no printable area)
            if (ex.Message != null && ex.Message.Contains("nothing was printed", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Save failed: No printable content was found in the workbook.");
                // Additional handling logic can be placed here (e.g., create a placeholder page, log, etc.)
            }
            else
            {
                // Re-throw or handle other CellsException types as needed
                Console.WriteLine($"An unexpected CellsException occurred: {ex.Message}");
                throw;
            }
        }
    }
}
