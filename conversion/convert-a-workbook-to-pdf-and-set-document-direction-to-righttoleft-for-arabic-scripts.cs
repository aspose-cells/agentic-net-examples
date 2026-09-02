// Title: Export an Excel workbook to PDF with right‑to‑left orientation for Arabic text using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an existing .xlsx file, sets Worksheet.DisplayRightToLeft to true, and saves the workbook as a PDF with Aspose.Cells. | Show how to enable right‑to‑left layout for Arabic scripts before converting a workbook to PDF using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# export Excel to PDF with right‑to‑left layout for Arabic | How to enable RTL display in worksheet before saving as PDF using Aspose.Cells | C# convert .xlsx to PDF with Arabic script direction using Aspose.Cells | Set DisplayRightToLeft property when saving workbook as PDF in .NET | Right‑to‑left PDF generation from Excel with Aspose.Cells library
// Tags: export workbook to PDF with RTL layout | DisplayRightToLeft property Aspose.Cells | Arabic script PDF conversion C# | right-to-left PDF generation using Aspose.Cells | SaveFormat.Pdf with RTL setting

using System;
using Aspose.Cells;

// The example loads or creates a workbook, sets the first worksheet's DisplayRightToLeft property to true to support Arabic scripts, and then saves the workbook as a PDF using Aspose.Cells.
class ConvertWorkbookToPdfRtl
{
    static void Main()
    {
        // Create a new workbook (or load an existing one by using new Workbook("input.xlsx"))
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable right‑to‑left display for Arabic scripts
        worksheet.DisplayRightToLeft = true;

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
