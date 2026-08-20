// Title: Export Japanese Era (Gengo) Dates to PDF using Aspose.Cells for .NET
// Description: This example creates a workbook, sets the region to Japan, inserts a date, applies a custom number format that shows the era (e.g., 令和2年9月15日), configures PdfSaveOptions with a Japanese‑capable font, and saves the file as a PDF so the era characters render correctly.
// Keywords: Aspose.Cells | Japanese era | gengo | PDF export | C# | .NET | regional settings Japan | custom number format | PdfSaveOptions | MS Gothic | localization | date formatting | Japanese font embedding
// Common Searches: Aspose.Cells display Japanese era in PDF | C# export workbook with gengo date format | set Japan region for date formatting Aspose.Cells | embed Japanese font in PDF using Aspose.Cells | custom number format ggge for Japanese dates | PDFSaveOptions Japanese characters missing glyphs
// Developer Intent: Generate a PDF from an Aspose.Cells workbook where dates are formatted with Japanese era symbols and displayed correctly.
// Use Cases: Financial statements for Japanese clients that require era‑based dates | Invoices or receipts complying with Japanese localization standards | Localized calendars or schedules exported as PDF with proper gengo representation | Regulatory reports in Japan where era notation is mandatory
// AI Prompts: Write C# code with Aspose.Cells to format a cell using the Japanese era pattern and export it to PDF, ensuring the characters are visible. | Explain how to set PdfSaveOptions.DefaultFont and CheckWorkbookDefaultFont to embed a Japanese font for correct era rendering. | Provide troubleshooting steps when era symbols appear as squares or missing glyphs in the generated PDF.

using System;
using Aspose.Cells;

// This example creates a workbook, sets the region to Japan, inserts a date, applies a custom number format that shows the era (e.g., 令和2年9月15日), configures PdfSaveOptions with a Japanese‑capable font, and saves the file as a PDF so the era characters render correctly.
class JapaneseEraPdfDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook's regional settings to Japan so that era formatting works
        workbook.Settings.Region = CountryCode.Japan;

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a sample date (e.g., 2020-09-15)
        sheet.Cells["A1"].PutValue(new DateTime(2020, 9, 15));

        // Apply a custom number format that displays the Japanese era (gengo)
        // Format example: "令和2年9月15日"
        Style style = sheet.Cells["A1"].GetStyle();
        style.Custom = "[$-ja-JP]ggge\"年\"m\"月\"d\"日\"";
        sheet.Cells["A1"].SetStyle(style);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use a font that contains Japanese characters to avoid missing glyphs
        pdfOptions.DefaultFont = "MS Gothic"; // or any other Japanese-capable font installed on the system
        pdfOptions.CheckWorkbookDefaultFont = true; // Try workbook default font first

        // Save the workbook as PDF
        workbook.Save("JapaneseEra.pdf", pdfOptions);
    }
}
