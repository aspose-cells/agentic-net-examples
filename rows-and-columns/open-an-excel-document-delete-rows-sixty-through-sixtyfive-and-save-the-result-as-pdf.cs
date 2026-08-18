// Title: C# – Delete rows 60‑65 in an Excel workbook and export to PDF with Aspose.Cells
// Description: Load an XLSX file using Aspose.Cells for .NET, remove rows 60‑65 from the first worksheet, and generate a PDF of the updated sheet. The example shows both a temporary XLSX save and direct PDF conversion via PdfSaveOptions or ConversionUtility.
// Keywords: Aspose.Cells C# | .NET Excel row deletion | delete rows 60-65 | Excel to PDF conversion | Workbook.Save PDF | PdfSaveOptions | ConversionUtility | programmatic Excel editing | remove specific rows Aspose | export worksheet to PDF
// Common Searches: Aspose.Cells delete rows 60 to 65 C# | Convert modified Excel to PDF using Aspose.Cells | How to remove specific rows from an Excel file in .NET | Aspose.Cells example for row deletion and PDF export | C# code to delete rows and save workbook as PDF
// Developer Intent: Programmatically strip rows 60‑65 from an Excel sheet and produce a PDF of the cleaned worksheet.
// Use Cases: Prepare financial reports by eliminating placeholder rows before publishing as PDF. | Automate data‑cleaning pipelines where certain rows must be removed prior to archiving. | Create printable PDF snapshots of worksheets after trimming excess rows for documentation.
// AI Prompts: Generate C# code that deletes rows 60‑65 with Aspose.Cells and saves directly to PDF without an intermediate file. | Add robust error handling for missing rows when using Workbook.Worksheets[0].Cells.DeleteRows. | Explain how to use PdfSaveOptions to control PDF output quality after row deletion.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Load an XLSX file using Aspose.Cells for .NET, remove rows 60‑65 from the first worksheet, and generate a PDF of the updated sheet. The example shows both a temporary XLSX save and direct PDF conversion via PdfSaveOptions or ConversionUtility.
class DeleteRowsAndConvertToPdf
{
    static void Main()
    {
        // Paths for the original Excel file, the intermediate modified file, and the final PDF.
        string inputExcelPath = "input.xlsx";
        string modifiedExcelPath = "modified.xlsx";
        string outputPdfPath = "output.pdf";

        // Load the existing workbook.
        Workbook workbook = new Workbook(inputExcelPath);

        // Delete rows 60 through 65 (zero‑based index: start at 59, delete 6 rows).
        workbook.Worksheets[0].Cells.DeleteRows(59, 6);

        // Save the modified workbook to a temporary XLSX file.
        workbook.Save(modifiedExcelPath, SaveFormat.Xlsx);

        // Convert the modified Excel file to PDF using the provided ConversionUtility rule.
        ConversionUtility.Convert(modifiedExcelPath, outputPdfPath);

        Console.WriteLine("Rows deleted and PDF saved successfully.");
    }
}
