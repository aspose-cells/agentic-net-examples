using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the TSV file into a workbook
        var loadOptions = new LoadOptions(LoadFormat.TabDelimited);
        Workbook workbook = new Workbook("input.tsv", loadOptions);

        // Configure PDF save options: one page per sheet, default (no) security
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true
            // SecurityOptions left unset to use default (no password protection)
        };

        // Export the workbook to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example – loads TSV, sets OnePagePerSheet, saves PDF with default security.