// Title: C# – Insert Column at Index 3 and Export Worksheet to PDF with Aspose.Cells for .NET
// Description: Loads an Excel workbook, inserts a new column at zero‑based index 3 while updating references, and saves the worksheet as a PDF. Includes file‑existence validation and exception handling.
// Keywords: Aspose.Cells | C# | .NET | insert column | Excel to PDF | worksheet column insertion | PDF export | code example | GitHub sample | Aspose.Cells PDFSaveOptions
// Common Searches: Aspose.Cells insert column C# example | Export Excel worksheet to PDF after adding column | C# insert column at index 3 using Aspose.Cells | How to save modified workbook as PDF with Aspose.Cells
// Developer Intent: Add a column at position three in an Excel sheet and generate a PDF version of the updated worksheet.
// Use Cases: Add a placeholder column before existing data, then create a printable PDF report. | Shift formulas by inserting a column and export financial statements to PDF with correct references. | Re‑format a worksheet layout by inserting a column before distributing a PDF to stakeholders.
// AI Prompts: Write C# code that uses Aspose.Cells to insert a column at index 3 and save the workbook as a PDF with custom page settings. | Provide a robust Aspose.Cells example that checks for the source Excel file, inserts a column, updates references, and handles errors while exporting to PDF.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, inserts a new column at zero‑based index 3 while updating references, and saves the worksheet as a PDF. Includes file‑existence validation and exception handling.
class Program
{
    static void Main()
    {
        // Paths for input Excel file and output PDF file
        string inputFile = "input.xlsx";
        string outputPdf = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a new column at index 3 (fourth column, zero‑based)
            // The second parameter updates references in other worksheets
            worksheet.Cells.InsertColumn(3, true);

            // Prepare PDF save options (optional, defaults are fine)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the modified workbook as a PDF file
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF: {outputPdf}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
