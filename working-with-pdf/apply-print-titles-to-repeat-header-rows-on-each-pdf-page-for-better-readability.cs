// Title: Repeat header rows on each PDF page with Aspose.Cells for .NET (C#)
// AI Prompts: Configure the worksheet's PageSetup.PrintTitleRows to the first row and save the workbook as a PDF using Aspose.Cells in C#. | Set print title rows (and optionally columns) before exporting an Excel file to PDF with Aspose.Cells for .NET. | Apply page‑setup settings to repeat a header row on every PDF page and generate the PDF via Workbook.Save.
// Common Searches: Aspose.Cells C# repeat first row as header on every PDF page | How to set PrintTitleRows in Aspose.Cells before PDF export | Export Excel to PDF with repeated header rows using Aspose.Cells .NET | PageSetup.PrintTitleRows example for PDF generation in C# | Aspose.Cells repeat column titles on each PDF page C#
// Tags: Aspose.Cells PrintTitleRows PDF export | C# page setup header row repetition | Export workbook to PDF with titles | Configure print title columns Aspose.Cells | Workbook.Save PDF with header repetition | Aspose.Cells page setup PDF conversion

using System;
using Aspose.Cells;

// The example loads an Excel workbook, sets the first row as a print title so it repeats on every printed page via the worksheet's PageSetup, and then saves the workbook as a PDF, preserving the repeated header.
class PrintTitlesToPdf
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Set the rows to repeat on each printed page (e.g., first row as header)
        // The range is specified in A1 style. "$1:$1" means row 1.
        sheet.PageSetup.PrintTitleRows = "$1:$1";

        // Optional: If you also want to repeat columns, set PrintTitleColumns similarly
        // sheet.PageSetup.PrintTitleColumns = "$A:$A";

        // Save the worksheet as a PDF. The print titles will be applied automatically.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
