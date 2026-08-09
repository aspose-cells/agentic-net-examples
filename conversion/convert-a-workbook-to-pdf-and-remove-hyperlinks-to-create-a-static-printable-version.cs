// Title: C# – Convert Excel to PDF and Strip All Hyperlinks with Aspose.Cells
// Description: Load an Excel workbook using Aspose.Cells for .NET, clear every worksheet's Hyperlinks collection, and save the result as a static, non‑interactive PDF ready for printing or archiving.
// Keywords: Aspose.Cells C# PDF conversion | remove hyperlinks Excel | Excel to PDF without links | clear worksheet hyperlinks | static printable PDF | Aspose.Cells SaveFormat.Pdf | batch Excel PDF conversion
// Common Searches: Aspose.Cells remove hyperlinks before PDF export | C# convert .xlsx to PDF without clickable links | how to clear all hyperlinks in an Excel workbook using Aspose | generate printable PDF from Excel with Aspose.Cells | strip hyperlinks from Excel and save as PDF C#
// Developer Intent: Produce a PDF version of an Excel file where all hyperlinks are removed, yielding a static document suitable for printing or compliance purposes.
// Use Cases: Creating printable reports from Excel templates that must not contain active links. | Preparing regulatory or legal documents by converting spreadsheets to PDF without clickable URLs. | Automating batch conversion of multiple workbooks to PDF while stripping hyperlinks for archival storage.
// AI Prompts: Write C# code with Aspose.Cells to load an .xlsx file, delete every hyperlink on all worksheets, and export it as a PDF. | Explain how Aspose.Cells SaveOptions can be configured to ensure hyperlinks are excluded during PDF export. | Provide a script that processes a directory of Excel files, removes all hyperlinks from each workbook, and saves them as PDFs using Aspose.Cells.

using System;
using Aspose.Cells;

// Load an Excel workbook using Aspose.Cells for .NET, clear every worksheet's Hyperlinks collection, and save the result as a static, non‑interactive PDF ready for printing or archiving.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath); // Load rule

        // Remove all hyperlinks from every worksheet to make the document static
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Hyperlinks.Clear(); // Clear hyperlink collection
        }

        // Save the workbook as a PDF file (static printable version)
        workbook.Save("output.pdf", SaveFormat.Pdf); // Save rule
    }
}
