using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExportDocumentStructureDemo
{
    static void Main()
    {
        // Create a new workbook with three worksheets
        Workbook workbook = new Workbook();
        Worksheet summarySheet = workbook.Worksheets[0];
        summarySheet.Name = "Summary";

        Worksheet detailsSheet = workbook.Worksheets.Add("Details");
        Worksheet dataSheet = workbook.Worksheets.Add("Data");

        // Fill each sheet with some sample data
        summarySheet.Cells["A1"].PutValue("Summary Sheet");
        detailsSheet.Cells["A1"].PutValue("Details Sheet");
        dataSheet.Cells["A1"].PutValue("Data Sheet");

        // Create an outline (grouped rows and columns) on the "Data" sheet
        // This will be reflected in the PDF outline when ExportDocumentStructure is true
        dataSheet.Cells.GroupRows(0, 4, true);      // Group rows 1‑5
        dataSheet.Cells.GroupColumns(0, 2, true);  // Group columns A‑C
        dataSheet.IsOutlineShown = true;           // Ensure the outline is visible in Excel

        // Configure PDF save options to export the document structure (outline)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true
        };

        // Save the workbook as a PDF; the resulting PDF will contain an outline
        // that mirrors the worksheet hierarchy and the outline created above
        string pdfPath = "WorkbookWithOutline.pdf";
        workbook.Save(pdfPath, pdfOptions);

        Console.WriteLine($"PDF saved to '{pdfPath}' with document structure exported.");
    }
}