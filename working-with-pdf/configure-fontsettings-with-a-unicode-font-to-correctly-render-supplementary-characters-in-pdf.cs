using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate the workbook with Unicode characters
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Unicode test: 😀 𝟘𝟙𝟚");
        sheet.Cells["A2"].PutValue("中文字符示例");

        // Configure PDF save options to use a Unicode‑capable default font
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DefaultFont = "MS Gothic",               // fallback font for supplementary characters
            CheckWorkbookDefaultFont = true,         // try workbook's default font first
            FontEncoding = PdfFontEncoding.Identity, // full Unicode support
            EmbedStandardWindowsFonts = true         // embed standard fonts for ASCII range
        };

        // Save the workbook as PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author: Example demonstrating FontSettings for Unicode PDF rendering.