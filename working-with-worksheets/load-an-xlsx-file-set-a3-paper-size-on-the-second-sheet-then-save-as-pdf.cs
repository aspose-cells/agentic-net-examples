// Title: C# – Set A3 Paper Size on the Second Worksheet and Convert XLSX to PDF with Aspose.Cells
// Description: Load an XLSX workbook, verify a second worksheet exists, change its PageSetup to A3, and save the entire workbook as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# set worksheet paper size | A3 page setup Aspose.Cells | Excel to PDF conversion .NET | modify second sheet page layout | Aspose.Cells SaveFormat.Pdf
// Common Searches: how to set A3 size for a specific sheet in Aspose.Cells | convert Excel workbook to PDF after changing page setup C# | check for second worksheet before applying page settings Aspose.Cells | C# code to export XLSX as PDF with custom paper size
// Developer Intent: Load an existing XLSX file, apply an A3 paper size to the second worksheet only, and export the workbook to a PDF document.
// Use Cases: Generate printable A3‑format reports from multi‑sheet Excel files. | Prepare invoices or statements on the second tab with A3 dimensions before PDF conversion. | Safely apply page‑setup changes only when a second worksheet is present to avoid runtime errors.
// AI Prompts: Write C# code that sets the paper size of the third worksheet to Letter and saves the workbook as PDF using Aspose.Cells. | Explain strategies for handling missing worksheets when modifying PageSetup properties in Aspose.Cells. | Provide a script to batch‑process a folder of Excel files, setting each sheet to A3 and converting them to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

// Load an XLSX workbook, verify a second worksheet exists, change its PageSetup to A3, and save the entire workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourceFile = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourceFile);

        // Verify that a second worksheet exists (index 1)
        if (workbook.Worksheets.Count > 1)
        {
            // Set the paper size of the second worksheet to A3
            workbook.Worksheets[1].PageSetup.PaperSize = PaperSizeType.PaperA3;
        }
        else
        {
            Console.WriteLine("The workbook does not contain a second worksheet.");
        }

        // Path for the resulting PDF file
        string pdfFile = "output.pdf";

        // Save the modified workbook as PDF
        workbook.Save(pdfFile, SaveFormat.Pdf);

        Console.WriteLine($"Workbook saved as PDF to '{pdfFile}'.");
    }
}
