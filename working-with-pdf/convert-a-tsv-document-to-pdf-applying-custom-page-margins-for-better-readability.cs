// Title: Convert a TSV file to PDF with custom page margins using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert TSV to PDF with specific left and right margins | How to set page margins before saving a workbook as PDF in Aspose.Cells | Aspose.Cells load tab‑separated values and export to PDF with custom margins | Set top and bottom margins for PDF export from TSV using Aspose.Cells .NET | Example code for TSV to PDF conversion with margin settings in C#
// Tags: Aspose.Cells TSV to PDF conversion | worksheet margin settings Aspose.Cells | set page margins before PDF export C# | LoadOptions TSV Aspose.Cells | export workbook as PDF with margins

using Aspose.Cells;
using System;

// The sample loads a TSV file into an Aspose.Cells Workbook, configures left, right, top, and bottom page margins on the first worksheet, and saves the workbook as a PDF document.
class Program
{
    static void Main()
    {
        // Load the TSV file into a workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook("input.tsv", loadOptions);

        // Access the first worksheet (or adjust as needed)
        Worksheet sheet = workbook.Worksheets[0];
        PageSetup pageSetup = sheet.PageSetup;

        // Set custom page margins (values are in inches)
        pageSetup.LeftMargin = 0.5;    // left margin
        pageSetup.RightMargin = 0.5;   // right margin
        pageSetup.TopMargin = 0.75;    // top margin
        pageSetup.BottomMargin = 0.75; // bottom margin

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
