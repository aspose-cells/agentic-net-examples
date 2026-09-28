// Title: Export an Excel workbook to PDF with English UI and document language for spell checking using Aspose.Cells in C#
// AI Prompts: Write C# code that saves a Workbook as a PDF while setting the workbook UI language to USA and the PDF default edit language to English with Aspose.Cells. | Show how to assign the built‑in document property Language to "en-US" and configure PdfSaveOptions for English spell checking before exporting to PDF.
// Common Searches: Aspose.Cells C# export workbook to PDF with English spell checking | how to set default edit language for PDF output in Aspose.Cells | C# set workbook UI language to USA before saving as PDF using Aspose | configure built‑in document property Language en-US in Aspose.Cells PDF conversion
// Tags: Aspose.Cells PDF export English language | C# set workbook UI language USA | PdfSaveOptions DefaultEditLanguage English | Workbook built‑in document property Language en-US | Excel to PDF conversion with spell checking

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsLanguagePdfDemo
{
    // The example creates a workbook, adds sample data, sets the UI language to USA, assigns the built‑in document property Language to "en-US", configures PdfSaveOptions.DefaultEditLanguage to English, and saves the workbook as a PDF, enabling proper English spell checking.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].PutValue("Sample text for spell checking.");

            // Set the workbook UI language to English (United States)
            workbook.Settings.LanguageCode = CountryCode.USA;

            // Set the built‑in document property "Language" to en‑US
            workbook.BuiltInDocumentProperties.Language = "en-US";

            // Configure PDF save options and set the default edit language to English
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.DefaultEditLanguage = DefaultEditLanguage.English;

            // Save the workbook as PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
