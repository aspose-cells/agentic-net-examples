// Title: C# – Remove Duplicate Rows (Preserve Formulas) and Export Excel to PDF with Aspose.Cells
// Description: Loads an XLSX file into an Aspose.Cells Workbook, eliminates repeated rows while keeping all formula cells intact, and saves the cleaned worksheet as a PDF document. Demonstrates the Worksheet.Cells.RemoveDuplicates method followed by PDF export.
// Keywords: Aspose.Cells | C# RemoveDuplicates | Excel duplicate rows | preserve formulas | Excel to PDF conversion | worksheet.Cells.RemoveDuplicates example | Aspose.Cells PDF export | deduplicate spreadsheet | batch Excel processing .NET | Aspose.Cells .NET
// Common Searches: Aspose.Cells remove duplicate rows C# | Keep formulas when deleting duplicate rows in Excel using Aspose | Export cleaned Excel workbook to PDF with Aspose.Cells | C# code to deduplicate worksheet and save as PDF | How to use RemoveDuplicates method in Aspose.Cells
// Developer Intent: The developer needs to load an Excel workbook, strip out repeated rows without breaking any formulas, and generate a PDF version of the resulting file.
// Use Cases: Sanitize sales reports by discarding repeated entries before creating a PDF summary for clients. | Automate the preparation of financial models where formula integrity must remain after data deduplication. | Process large batches of uploaded spreadsheets, remove redundancies, and deliver PDF snapshots for archival.
// AI Prompts: Write C# code that uses Aspose.Cells to delete duplicate rows in a specific worksheet, ensures all formula cells stay functional, and then saves the workbook as a PDF. | Show how to configure the RemoveDuplicates call to target selected columns, retain formula calculations, and export the cleaned sheet to PDF in a single Aspose.Cells workflow.

using System;
using Aspose.Cells;

namespace AsposeCellsDuplicateRemoval
{
    // Loads an XLSX file into an Aspose.Cells Workbook, eliminates repeated rows while keeping all formula cells intact, and saves the cleaned worksheet as a PDF document. Demonstrates the Worksheet.Cells.RemoveDuplicates method followed by PDF export.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Path for the resulting PDF file
            string pdfFile = "output.pdf";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(sourceFile);

            // Access the first worksheet (you can change the index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Remove duplicate rows in the worksheet while keeping formulas intact
            worksheet.Cells.RemoveDuplicates();

            // Save the cleaned workbook as a PDF document
            workbook.Save(pdfFile, SaveFormat.Pdf);

            Console.WriteLine("Duplicate rows removed and PDF saved successfully.");
        }
    }
}
