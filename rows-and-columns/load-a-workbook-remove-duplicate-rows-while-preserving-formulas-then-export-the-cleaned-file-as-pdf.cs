// Title: C# – Remove Duplicate Rows (Keep Formulas) and Convert Excel to PDF with Aspose.Cells
// Description: Loads an Excel workbook, applies Worksheet.Cells.RemoveDuplicates to purge repeated rows while preserving any formulas, saves the sanitized file, and then creates a PDF using ConversionUtility.Convert.
// Keywords: Aspose.Cells C# duplicate rows | preserve formulas Aspose.Cells | Excel to PDF conversion .NET | Worksheet.RemoveDuplicates example | ConversionUtility.Convert PDF | programmatic data deduplication Excel | Aspose.Cells data cleansing
// Common Searches: Aspose.Cells remove duplicate rows C# | Keep formulas when deleting duplicate rows in Excel with Aspose | Convert cleaned Excel workbook to PDF using Aspose.Cells .NET | How to deduplicate first worksheet with Aspose.Cells | Batch process Excel files to eliminate duplicates and export PDFs
// Developer Intent: Strip repeated rows from a worksheet without breaking formulas and output the result as a PDF document.
// Use Cases: Prepare client‑ready reports by removing redundant entries before PDF distribution. | Automate monthly financial data cleanup, ensuring formulas stay functional, then archive as PDF. | Standardize large data sets for regulatory filing by deduplicating and converting to a non‑editable format.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete duplicate rows while retaining formulas and then saves the workbook as a PDF. | Explain the behavior of Worksheet.Cells.RemoveDuplicates with formula‑containing cells and how ConversionUtility.Convert creates the PDF output. | Recommend performance‑optimizing techniques for processing massive Excel files when performing deduplication and PDF conversion with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsDuplicateRemoval
{
    // Loads an Excel workbook, applies Worksheet.Cells.RemoveDuplicates to purge repeated rows while preserving any formulas, saves the sanitized file, and then creates a PDF using ConversionUtility.Convert.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string inputPath = "input.xlsx";

            // Load the workbook (uses the Workbook(string) constructor rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Remove duplicate rows while keeping formulas intact
            // (uses Cells.RemoveDuplicates() rule)
            worksheet.Cells.RemoveDuplicates();

            // Save the cleaned workbook to a temporary file (uses Workbook.Save(string) rule)
            string cleanedPath = "cleaned.xlsx";
            workbook.Save(cleanedPath);

            // Convert the cleaned Excel file to PDF (uses ConversionUtility.Convert(string, string) rule)
            string pdfPath = "output.pdf";
            ConversionUtility.Convert(cleanedPath, pdfPath);

            Console.WriteLine("Duplicate rows removed and PDF generated successfully.");
        }
    }
}
