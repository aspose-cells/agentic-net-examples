// Title: Convert Excel to PDF and strip hyperlinks with Aspose.Cells (C#)
// Description: Load an Excel workbook, clear every worksheet's Hyperlinks collection, and save it as a PDF using Aspose.Cells for .NET to create a static, print‑ready document.
// Keywords: Aspose.Cells | C# | Excel to PDF | remove hyperlinks | clear worksheet hyperlinks | static PDF | printable PDF | Workbook.Save | SaveFormat.Pdf
// Common Searches: Aspose.Cells remove hyperlinks before PDF export | C# convert Excel to PDF without links | how to clear hyperlinks in Excel using Aspose.Cells | export Excel as printable PDF .NET | strip hyperlinks from workbook then save as PDF
// Developer Intent: Strip all hyperlinks from an Excel workbook and export it as a PDF.
// Use Cases: Generate a print‑ready PDF report from an Excel template that contains hyperlink formulas. | Archive financial spreadsheets as PDFs where clickable links must be disabled. | Distribute static documentation PDFs from Excel files without interactive elements.
// AI Prompts: Write C# code with Aspose.Cells that loads an Excel file, removes every hyperlink, and saves the workbook as a PDF. | Explain the steps to clear the Hyperlinks collection for each worksheet before PDF conversion using Aspose.Cells, and note any performance tips. | Provide a concise tutorial for converting an Excel workbook to a printable PDF while ensuring no hyperlinks remain, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an Excel workbook, clear every worksheet's Hyperlinks collection, and save it as a PDF using Aspose.Cells for .NET to create a static, print‑ready document.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourceFile = "input.xlsx";

        // Load the workbook from the file (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(sourceFile);

        // Iterate through all worksheets and clear their Hyperlinks collections
        // This removes all hyperlinks, making the document static for printing
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Hyperlinks.Clear();
        }

        // Path for the resulting PDF file
        string pdfFile = "output.pdf";

        // Save the modified workbook as PDF (uses Workbook.Save(string, SaveFormat))
        workbook.Save(pdfFile, SaveFormat.Pdf);
    }
}
