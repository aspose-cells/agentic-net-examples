// Title: Remove all charts from an Excel workbook and save it as a PDF while preserving original cell formatting with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, deletes every chart from each worksheet, and then saves the workbook as a PDF keeping all cell styles unchanged. | Show how to iterate through the Worksheets collection, clear the Charts collection, and call Workbook.Save with SaveFormat.Pdf to produce a formatted PDF. | Provide a snippet that conditionally skips chart removal on a named sheet but still exports the entire workbook to PDF with formatting retained.
// Common Searches: Aspose.Cells C# remove charts from all worksheets before PDF export | How to export Excel to PDF without charts while keeping cell formatting in .NET | Clear chart objects in Aspose.Cells and generate PDF preserving styles | Save workbook as PDF with original formatting after deleting charts using Aspose.Cells | C# code to delete Excel charts and convert to PDF with Aspose.Cells
// Tags: clear worksheet charts Aspose.Cells | export workbook to PDF preserving formatting | remove Excel charts before PDF conversion .NET | Aspose.Cells chart collection clear | save workbook as PDF without charts

using Aspose.Cells;
using System;

// // Loads an Excel file, clears all chart objects from each worksheet, and saves the workbook as a PDF while retaining the original cell formatting.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        string workbookPath = "input.xlsx";          // path to the source workbook
        Workbook workbook = new Workbook(workbookPath);

        // Remove all charts from every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Clear the Charts collection to delete all charts on the sheet
            sheet.Charts.Clear();
        }

        // Save the modified workbook as a PDF while preserving original cell formatting
        string pdfPath = "output.pdf";               // desired PDF output path
        workbook.Save(pdfPath, SaveFormat.Pdf);
    }
}
