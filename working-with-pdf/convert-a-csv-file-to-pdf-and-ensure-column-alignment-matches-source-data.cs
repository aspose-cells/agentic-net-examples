// Title: C# – Convert CSV to PDF with exact column alignment using Aspose.Cells
// Description: This example shows how to load a CSV file, convert it to an intermediate XLSX workbook, auto‑fit the columns, and save the result as a PDF/A‑1b document, ensuring the PDF layout matches the original CSV column widths.
// Keywords: Aspose.Cells CSV to PDF | C# auto fit columns PDF | preserve column widths Aspose | PDF/A-1b export .NET | CSV to PDF conversion example | Aspose.Cells workbook to PDF | ConversionUtility CSV XLSX
// Common Searches: how to keep CSV column widths when exporting to PDF with Aspose.Cells | Aspose.Cells C# convert CSV to PDF with PDF/A compliance | auto fit columns before saving workbook as PDF Aspose | preserve formatting CSV to PDF .NET | batch convert CSV files to PDF using Aspose.Cells
// Developer Intent: Create a PDF from a CSV file while retaining the original column layout.
// Use Cases: Generate printable reports from CSV exports with exact column alignment. | Produce archival‑ready PDF/A‑1b files from data tables without losing formatting. | Automate large‑scale conversion of CSV datasets to PDFs with consistent layout.
// AI Prompts: Write C# code that uses Aspose.Cells to convert a CSV file to PDF, auto‑fit columns, and apply PDF/A‑1b compliance. | Explain the role of ConversionUtility.Convert and Worksheet.AutoFitColumns in preserving column alignment during PDF export. | Provide a step‑by‑step guide for batch processing a folder of CSV files into PDFs while maintaining formatting with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

// This example shows how to load a CSV file, convert it to an intermediate XLSX workbook, auto‑fit the columns, and save the result as a PDF/A‑1b document, ensuring the PDF layout matches the original CSV column widths.
class CsvToPdfConverter
{
    static void Main()
    {
        // Paths for source CSV and intermediate XLSX
        string csvPath = "input.csv";
        string xlsxPath = "intermediate.xlsx";
        string pdfPath = "output.pdf";

        // 1. Convert CSV to XLSX using the provided ConversionUtility rule
        // This creates a workbook that matches the CSV data layout.
        ConversionUtility.Convert(csvPath, xlsxPath);

        // 2. Load the generated XLSX workbook
        Workbook workbook = new Workbook(xlsxPath);
        Worksheet worksheet = workbook.Worksheets[0];

        // 3. Adjust column widths to fit the imported data.
        // This ensures the column alignment in the PDF mirrors the source CSV.
        worksheet.AutoFitColumns();

        // 4. Set PDF save options (optional: set compliance level)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA1b; // keep PDF/A-1b compliance

        // 5. Save the workbook as PDF using the provided Save method.
        workbook.Save(pdfPath, pdfOptions);

        Console.WriteLine($"CSV file '{csvPath}' has been converted to PDF '{pdfPath}' with column alignment preserved.");
    }
}
