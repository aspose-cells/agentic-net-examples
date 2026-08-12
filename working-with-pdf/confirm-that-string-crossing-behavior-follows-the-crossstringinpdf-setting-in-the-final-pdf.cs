// Title: Aspose.Cells .NET: Verify TextCrossType (CrossKeep, CrossOverride, StrictInCell) in PDF Export
// Description: C# sample that writes a long string to cell A1, narrows column A, and saves three PDFs using PdfSaveOptions with TextCrossType set to CrossKeep, CrossOverride, and StrictInCell. Demonstrates how each setting controls text overflow and cell content handling in the generated PDF.
// Keywords: Aspose.Cells | .NET | C# | PdfSaveOptions | TextCrossType | CrossKeep | CrossOverride | StrictInCell | PDF export | text overflow | cell boundary | Aspose.Cells example | GitHub | coding tutorial
// Common Searches: Aspose.Cells TextCrossType PDF example | How does CrossKeep affect PDF text overflow | CrossOverride vs StrictInCell in Aspose.Cells PDF | Verify text crossing setting in generated PDF | C# Aspose.Cells PDF export overflow text
// Developer Intent: Ensure the PDF output follows the selected TextCrossType option for handling overflowing cell text.
// Use Cases: Export a PDF where overflow text keeps original cell data (CrossKeep). | Export a PDF where overflow text overwrites adjacent cells (CrossOverride). | Export a PDF that truncates overflow text to stay inside the cell (StrictInCell).
// AI Prompts: Create a unit test that opens CrossKeep.pdf, CrossOverride.pdf, and StrictInCell.pdf and asserts the presence or absence of text in neighboring cells according to each TextCrossType. | Write a step‑by‑step guide to compare the visual rendering of the three PDFs using Aspose.PDF, highlighting differences caused by TextCrossType settings. | Generate a CI script that runs the Aspose.Cells code, produces the PDFs, and automatically validates text crossing behavior with a PDF parsing library.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCrossStringDemo
{
    // C# sample that writes a long string to cell A1, narrows column A, and saves three PDFs using PdfSaveOptions with TextCrossType set to CrossKeep, CrossOverride, and StrictInCell. Demonstrates how each setting controls text overflow and cell content handling in the generated PDF.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a long text in A1 that will exceed the column width
            sheet.Cells["A1"].PutValue("This is a very long text that will definitely cross the cell boundary when the column is narrow.");

            // Set a narrow column width to force overflow
            sheet.Cells.SetColumnWidth(0, 5); // column A width = 5 characters

            // ---------- Demonstrate TextCrossType.CrossKeep ----------
            PdfSaveOptions optionsCrossKeep = new PdfSaveOptions();
            optionsCrossKeep.TextCrossType = TextCrossType.CrossKeep; // text will cross cells and keep existing cell contents
            workbook.Save("CrossKeep.pdf", optionsCrossKeep);
            Console.WriteLine("PDF saved with TextCrossType.CrossKeep");

            // ---------- Demonstrate TextCrossType.CrossOverride ----------
            PdfSaveOptions optionsCrossOverride = new PdfSaveOptions();
            optionsCrossOverride.TextCrossType = TextCrossType.CrossOverride; // text will cross cells and override existing cell contents
            workbook.Save("CrossOverride.pdf", optionsCrossOverride);
            Console.WriteLine("PDF saved with TextCrossType.CrossOverride");

            // ---------- Demonstrate TextCrossType.StrictInCell ----------
            PdfSaveOptions optionsStrict = new PdfSaveOptions();
            optionsStrict.TextCrossType = TextCrossType.StrictInCell; // text will be truncated to stay within the cell
            workbook.Save("StrictInCell.pdf", optionsStrict);
            Console.WriteLine("PDF saved with TextCrossType.StrictInCell");
        }
    }
}
