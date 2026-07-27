using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Author: Aspose.Cells .NET example – set default font for Unicode characters when saving to PDF
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample Unicode text that may not have a proper font in the cell style
        sheet.Cells["A1"].PutValue("漢字テスト"); // Japanese characters
        sheet.Cells["A2"].PutValue("Пример текста"); // Cyrillic characters

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set a font that supports the required Unicode ranges (e.g., MS Gothic, MingLiu)
        pdfOptions.DefaultFont = "MS Gothic";

        // Try to use the workbook's default font first, then fall back to DefaultFont
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}