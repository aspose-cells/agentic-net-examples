// Title: C# – Remove Empty Columns from an Excel Sheet and Save as PDF using Aspose.Cells
// Description: Loads an XLSX workbook, eliminates all columns without data in the first worksheet via DeleteBlankColumns, and converts the cleaned sheet to a PDF file using the Save method with SaveFormat.Pdf.
// Keywords: Aspose.Cells C# delete empty columns | Excel to PDF conversion .NET | Remove blank columns Aspose | DeleteBlankColumns method example | SaveFormat.Pdf usage | trim worksheet columns C# | Aspose.Cells PDF export
// Common Searches: how to delete all empty columns in an Excel file with Aspose.Cells C# | convert a cleaned Excel worksheet to PDF using Aspose.Cells .NET | Aspose.Cells DeleteBlankColumns first worksheet example | C# code to trim Excel columns and export to PDF | remove blank columns before PDF generation Aspose
// Developer Intent: Strip every column that contains no data from the primary worksheet and produce a PDF of the resulting sheet.
// Use Cases: Prepare a printable report by discarding unused columns and exporting the result as a compact PDF. | Automate data‑preprocessing pipelines where empty columns must be removed before archiving the spreadsheet as PDF. | Reduce file size for client‑facing documents by cleaning the worksheet and generating a PDF in one step.
// AI Prompts: Generate C# code that removes blank columns from all worksheets in a workbook and saves each cleaned sheet as an individual PDF. | Explain how DeleteBlankColumns treats merged cells and recommend techniques to keep original formatting when converting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an XLSX workbook, eliminates all columns without data in the first worksheet via DeleteBlankColumns, and converts the cleaned sheet to a PDF file using the Save method with SaveFormat.Pdf.
class Program
{
    static void Main()
    {
        // Load the source Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Delete all blank columns in the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.DeleteBlankColumns();

        // Save the trimmed worksheet as a PDF document
        workbook.Save("trimmed_output.pdf", SaveFormat.Pdf);
    }
}
