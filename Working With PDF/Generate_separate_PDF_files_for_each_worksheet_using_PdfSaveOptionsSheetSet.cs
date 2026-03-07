using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class SeparatePdfPerWorksheet
{
    static void Main()
    {
        // Create a new workbook and add a few worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Fill each worksheet with some sample data
        for (int s = 0; s < workbook.Worksheets.Count; s++)
        {
            Worksheet ws = workbook.Worksheets[s];
            ws.Cells["A1"].PutValue($"Data from {ws.Name}");
        }

        // Save each worksheet as an individual PDF file
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Configure SheetSet to include only the current sheet (0‑based index)
            pdfOptions.SheetSet = new SheetSet(new int[] { i });

            string outputFile = $"Worksheet_{i + 1}.pdf";
            workbook.Save(outputFile, pdfOptions);
        }
    }
}