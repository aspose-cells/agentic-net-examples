// Title: C# – Delete Row 8 from an Excel Workbook and Export to PDF with Aspose.Cells
// Description: Load an existing .xlsx file, remove the eighth row (index 7) from the first worksheet using Aspose.Cells for .NET, and save the modified workbook directly as a PDF document.
// Keywords: Aspose.Cells delete row C# | remove Excel row .NET | export Excel to PDF Aspose | row index 7 Aspose.Cells | C# Excel PDF conversion example
// Common Searches: how to delete a specific row in Excel using Aspose.Cells | Aspose.Cells C# export edited workbook to PDF | remove row 8 before PDF conversion Aspose | C# code to delete Excel row and save as PDF
// Developer Intent: Delete a targeted row from an Excel sheet and generate a PDF version of the updated file.
// Use Cases: Cleaning uploaded spreadsheets by stripping out header or placeholder rows prior to archiving as PDF. | Automating report pipelines where certain rows must be omitted before distribution. | Batch processing of Excel files to remove unwanted rows and produce PDF outputs for compliance.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete row 8 from the first worksheet of an .xlsx file and save the result as a PDF. | Explain Aspose.Cells row indexing and how to ensure the PDF reflects row deletions. | Add robust error handling for missing files, out‑of‑range row indices, and PDF save failures in the Aspose.Cells workflow.

using System;
using Aspose.Cells;

// Load an existing .xlsx file, remove the eighth row (index 7) from the first worksheet using Aspose.Cells for .NET, and save the modified workbook directly as a PDF document.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourceFile = "input.xlsx";

        // Path where the PDF will be saved
        string pdfFile = "output.pdf";

        // Load the workbook from the existing Excel file
        Workbook workbook = new Workbook(sourceFile);

        // Delete row 8 (zero‑based index is 7) from the first worksheet
        workbook.Worksheets[0].Cells.DeleteRow(7);

        // Save the modified workbook as a PDF document
        workbook.Save(pdfFile, SaveFormat.Pdf);
    }
}
