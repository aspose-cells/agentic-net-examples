// Title: C# – Remove Duplicate Rows (Preserve Formulas) and Export to PDF with Aspose.Cells
// Description: Loads an Excel workbook, removes duplicate rows while keeping all formulas intact using Worksheet.Cells.RemoveDuplicates, and saves the result as a PDF. A concise end‑to‑end example for Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# remove duplicate rows | Cells.RemoveDuplicates | preserve formulas | export Excel to PDF | Aspose.Cells PDF conversion | duplicate row removal .NET | Excel data cleaning Aspose
// Common Searches: How to delete duplicate rows in Excel using Aspose.Cells C# | Aspose.Cells keep formulas when removing duplicates | Convert cleaned Excel worksheet to PDF with Aspose.Cells | Remove duplicate rows and export to PDF in .NET | Worksheet.Cells.RemoveDuplicates example
// Developer Intent: Load an Excel file, eliminate duplicate rows without affecting any formulas, and generate a PDF version of the cleaned worksheet.
// Use Cases: Prepare a sales report by stripping duplicate entries, retaining calculated columns, and delivering a read‑only PDF to management. | Automate monthly financial statements: deduplicate data, keep all formula‑driven totals, and export the sheet as a PDF for stakeholder review. | Create archival copies of spreadsheets where duplicate rows are removed, formulas remain for reference, and the final document is saved as a PDF.
// AI Prompts: Generate C# code with Aspose.Cells that removes duplicate rows while preserving formulas and then saves the workbook as a PDF. | Explain the behavior of Worksheet.Cells.RemoveDuplicates in Aspose.Cells regarding formula retention and column comparison. | Add comprehensive error handling to a C# routine that loads an Excel file, removes duplicates, and exports the result to PDF using Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook, removes duplicate rows while keeping all formulas intact using Worksheet.Cells.RemoveDuplicates, and saves the result as a PDF. A concise end‑to‑end example for Aspose.Cells for .NET.
class RemoveDuplicatesAndExportPdf
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path for the resulting PDF file
        string pdfPath = "output.pdf";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourcePath);

        // Access the first worksheet (you can change the index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Remove duplicate rows in the worksheet while keeping formulas intact
        sheet.Cells.RemoveDuplicates();

        // Save the cleaned workbook as a PDF document
        workbook.Save(pdfPath, SaveFormat.Pdf);
    }
}
