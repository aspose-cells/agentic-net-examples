// Title: Insert three rows at row 20 while preserving formatting and export the worksheet to PDF using Aspose.Cells for .NET
// AI Prompts: Add three new rows starting at row 20 in an existing Excel file, copy the style from the row above, and generate a PDF output with Aspose.Cells in C#. | Show how to set InsertOptions to retain row formatting and then save the workbook as a PDF via PdfSaveOptions.
// Common Searches: Aspose.Cells insert rows at specific index keep original style C# | How to add rows and maintain formatting before converting Excel to PDF with Aspose.Cells | C# code to insert multiple rows and export workbook to PDF using Aspose.Cells | Preserve cell formatting when inserting rows in Aspose.Cells and save as PDF | InsertRows method with CopyFormatType SameAsAbove example
// Tags: InsertRows with CopyFormatType SameAsAbove Aspose.Cells | preserve row formatting during insertion C# | PdfSaveOptions for Excel to PDF conversion Aspose.Cells | Aspose.Cells InsertOptions usage example | export modified worksheet to PDF Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing; // for PdfSaveOptions if needed

// // Loads an existing Excel workbook, inserts three rows at row 20 copying the formatting from the preceding row, and saves the updated sheet as a PDF using Aspose.Cells.
class InsertRowsAndSavePdf
{
    static void Main()
    {
        // Paths for the input Excel file and the output PDF file
        string inputFile = "input.xlsx";
        string outputPdf = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare insert options to preserve formatting (copy format from the row above)
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove,
            UpdateReference = true
        };

        // Insert three rows starting at row 20 (zero‑based index 19)
        worksheet.Cells.InsertRows(19, 3, insertOptions);

        // Save the modified workbook as a PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save(outputPdf, pdfOptions);
    }
}
