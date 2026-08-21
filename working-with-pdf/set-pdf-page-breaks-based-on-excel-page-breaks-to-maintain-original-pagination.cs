// Title: C# – Convert Excel to PDF with Preserved Horizontal & Vertical Page Breaks using Aspose.Cells
// Description: Shows how to add custom horizontal and vertical page breaks in an Aspose.Cells workbook and export the worksheet to PDF while retaining the original pagination. The example uses PdfSaveOptions with default settings that keep the page‑break layout intact.
// Keywords: Aspose.Cells | C# | PDF conversion | Excel page breaks | horizontal page break | vertical page break | PdfSaveOptions | preserve pagination | worksheet to PDF | Aspose.Cells .NET
// Common Searches: Aspose.Cells keep page breaks when saving to PDF | C# export Excel to PDF with custom page breaks | preserve Excel pagination in PDF using Aspose | PdfSaveOptions page break settings | convert worksheet to PDF with row and column breaks
// Developer Intent: Export an Excel worksheet to PDF while maintaining any manually defined horizontal and vertical page breaks.
// Use Cases: Generate multi‑page PDF reports that follow the same section breaks as the source spreadsheet. | Create printable PDFs that match the layout of a workbook containing both row and column page breaks. | Automate batch conversion of Excel sheets to PDFs, ensuring the output pagination mirrors the original Excel design.
// AI Prompts: Write C# code with Aspose.Cells to add horizontal and vertical page breaks and save the worksheet as a PDF that preserves pagination. | Explain which PdfSaveOptions properties affect page‑break retention when converting Excel to PDF in Aspose.Cells for .NET. | Show how to retrieve automatic page‑break areas from a worksheet, compare them with custom breaks, and then export to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to add custom horizontal and vertical page breaks in an Aspose.Cells workbook and export the worksheet to PDF while retaining the original pagination. The example uses PdfSaveOptions with default settings that keep the page‑break layout intact.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data (100 rows)
        for (int i = 0; i < 100; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Add horizontal page breaks at rows 20 and 50 (zero‑based indices)
        worksheet.HorizontalPageBreaks.Add(19); // after row 20
        worksheet.HorizontalPageBreaks.Add(49); // after row 50

        // Add a vertical page break after column 5 (zero‑based index)
        worksheet.VerticalPageBreaks.Add(5);

        // Retrieve automatic page break areas (optional verification)
        ImageOrPrintOptions printOptions = new ImageOrPrintOptions();
        CellArea[] automaticBreaks = worksheet.GetPrintingPageBreaks(printOptions);
        Console.WriteLine($"Automatic page break areas detected: {automaticBreaks.Length}");

        // Save the workbook to PDF while preserving the defined page breaks
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Default settings keep pagination; no need to set OnePagePerSheet or AllColumnsInOnePagePerSheet
        workbook.Save("OutputWithPageBreaks.pdf", pdfOptions);
    }
}
