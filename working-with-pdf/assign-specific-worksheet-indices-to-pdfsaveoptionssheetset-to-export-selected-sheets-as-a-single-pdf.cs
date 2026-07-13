using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook with three worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Populate each sheet with sample data
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            sheet.Cells["A1"].PutValue($"Data in {sheet.Name}");
        }

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export only the first and third sheets (0‑based indexes)
        pdfOptions.SheetSet = new SheetSet(new int[] { 0, 2 });

        // Save the selected sheets as a single PDF
        workbook.Save("SelectedSheets.pdf", pdfOptions);
    }
}
// Author: Generated example demonstrating PdfSaveOptions.SheetSet usage.