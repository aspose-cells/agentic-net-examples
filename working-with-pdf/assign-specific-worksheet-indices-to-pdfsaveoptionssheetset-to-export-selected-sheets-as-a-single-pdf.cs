// Title: Export Selected Worksheets to a Single PDF with PdfSaveOptions.SheetSet (Aspose.Cells for .NET)
// Description: Shows how to set PdfSaveOptions.SheetSet with specific worksheet indexes (e.g., 0 and 2) to save only those sheets from an Aspose.Cells workbook as one PDF file using C#.
// Keywords: Aspose.Cells | PdfSaveOptions | SheetSet | export selected worksheets to PDF | C# Aspose.Cells PDF | worksheet index PDF export | .NET Excel to PDF | multiple sheets PDF Aspose | select sheets for PDF | Aspose.Cells PDF options
// Common Searches: Aspose.Cells export only certain worksheets to PDF | PdfSaveOptions SheetSet example C# | how to save specific Excel sheets as one PDF with Aspose | C# set sheet indexes in PdfSaveOptions | combine selected worksheets into a PDF Aspose.Cells
// Developer Intent: Generate a single PDF that contains only the worksheets you specify.
// Use Cases: Produce a client‑ready report that includes only the summary and conclusion tabs, leaving internal data sheets out. | Create an invoice PDF that merges the invoice details and terms worksheets while excluding auxiliary calculations. | Deliver a project overview PDF for stakeholders by selecting only the overview and milestones sheets from a larger workbook.
// AI Prompts: Write C# code that uses PdfSaveOptions.SheetSet to export worksheets 2 and 4 into one PDF with Aspose.Cells. | Show how to map sheet names to their indexes and pass them to SheetSet for PDF export in Aspose.Cells. | Explain how to combine selected worksheets into a PDF and set additional options like page orientation and compression using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to set PdfSaveOptions.SheetSet with specific worksheet indexes (e.g., 0 and 2) to save only those sheets from an Aspose.Cells workbook as one PDF file using C#.
class ExportSelectedSheetsToPdf
{
    static void Main()
    {
        // Create a new workbook with three worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "First";
        workbook.Worksheets.Add("Second");
        workbook.Worksheets.Add("Third");

        // Populate each sheet with sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Data in First sheet");
        workbook.Worksheets[1].Cells["A1"].PutValue("Data in Second sheet");
        workbook.Worksheets[2].Cells["A1"].PutValue("Data in Third sheet");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export only the first and third sheets (0‑based indexes 0 and 2)
        pdfOptions.SheetSet = new SheetSet(new int[] { 0, 2 });

        // Save the selected sheets as a single PDF file
        workbook.Save("SelectedSheets.pdf", pdfOptions);
    }
}
