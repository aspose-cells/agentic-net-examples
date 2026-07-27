// Title: Aspose.Cells .NET: Verify HtmlSaveOptions.WidthScalable Affects Only HTML Export
// Description: C# sample that creates a workbook, sets column widths, enables HtmlSaveOptions.WidthScalable, and saves to HTML, PDF, and XLSX. Demonstrates that scalable column widths apply only to HTML while PDF and XLSX files stay unchanged.
// Keywords: Aspose.Cells | .NET | C# | HtmlSaveOptions | WidthScalable | HTML export | PDF export | XLSX export | column width scaling | responsive HTML | non‑HTML formats | save options testing
// Common Searches: Does HtmlSaveOptions.WidthScalable affect PDF output in Aspose.Cells? | How to test WidthScalable only works for HTML export in C# | Aspose.Cells HTML export with scalable column widths | WidthScalable property impact on XLSX files | Aspose.Cells save workbook as HTML, PDF, XLSX simultaneously
// Developer Intent: Confirm that setting HtmlSaveOptions.WidthScalable changes column rendering only for HTML files and leaves PDF and XLSX exports untouched.
// Use Cases: Create responsive HTML reports while preserving original column sizes in PDF and XLSX. | Add an automated regression test to ensure WidthScalable does not alter non‑HTML save options. | Build a multi‑format export pipeline where HTML uses scalable widths for web display and other formats use default layout.
// AI Prompts: Generate a C# NUnit test that verifies column widths in PDF and XLSX remain unchanged when HtmlSaveOptions.WidthScalable is true. | Provide code to compare rendered column widths between HTML (WidthScalable enabled) and PDF outputs using Aspose.Cells. | Explain how to configure HtmlSaveOptions.WidthScalable for responsive HTML while keeping PDF and XLSX exports unaffected.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// C# sample that creates a workbook, sets column widths, enables HtmlSaveOptions.WidthScalable, and saves to HTML, PDF, and XLSX. Demonstrates that scalable column widths apply only to HTML while PDF and XLSX files stay unchanged.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Test WidthScalable Property");
        worksheet.Cells["B1"].PutValue(123.45);
        // Set column widths so the effect of WidthScalable can be observed in HTML
        worksheet.Cells.SetColumnWidth(0, 20);
        worksheet.Cells.SetColumnWidth(1, 20);

        // Configure HTML save options with WidthScalable enabled
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.WidthScalable = true; // Enable scalable column width for HTML

        // Save the workbook as HTML (WidthScalable should take effect here)
        workbook.Save("output_scalable.html", htmlOptions);

        // Save the same workbook as PDF – WidthScalable is irrelevant for PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("output.pdf", pdfOptions);

        // Save the workbook as XLSX – again, WidthScalable has no impact
        workbook.Save("output.xlsx");

        Console.WriteLine("Saved HTML, PDF, and XLSX files. WidthScalable only influences HTML output.");
    }
}
