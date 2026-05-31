using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class SeparatePdfPerWorksheet
{
    static void Main()
    {
        // Create a new workbook with three worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Populate each sheet with simple data
        for (int i = 0; i < 3; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            ws.Cells["A1"].PutValue($"Data for {ws.Name}");
            ws.Cells["A2"].PutValue(i + 1);
        }

        // Define distinct custom paper sizes (width, height in inches) for each sheet
        // Sheet1: 4" x 6"
        workbook.Worksheets[0].PageSetup.CustomPaperSize(4.0, 6.0);
        // Sheet2: 5" x 7"
        workbook.Worksheets[1].PageSetup.CustomPaperSize(5.0, 7.0);
        // Sheet3: 8.5" x 11" (standard Letter size)
        workbook.Worksheets[2].PageSetup.CustomPaperSize(8.5, 11.0);

        // Save each worksheet as a separate PDF using PdfSaveOptions and SheetSet
        for (int i = 0; i < 3; i++)
        {
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Restrict rendering to the current worksheet only
            pdfOptions.SheetSet = new SheetSet(new int[] { i });

            string pdfFileName = $"{workbook.Worksheets[i].Name}.pdf";
            workbook.Save(pdfFileName, pdfOptions);
        }
    }
}