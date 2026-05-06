using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample worksheets
        Workbook workbook = new Workbook();

        // First sheet (default)
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets[0].Cells["A1"].PutValue("Data for Sheet1");

        // Additional sheets
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].PutValue("Data for Sheet2");

        workbook.Worksheets.Add("Sheet3");
        workbook.Worksheets[2].Cells["A1"].PutValue("Data for Sheet3");

        // Iterate through each worksheet and save it as an individual PDF
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            // Configure PDF save options to include only the current sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.SheetSet = new SheetSet(new int[] { i }); // 0‑based index of the sheet

            // Build a file name that reflects the sheet index and name
            string pdfFileName = $"Worksheet_{i + 1}_{workbook.Worksheets[i].Name}.pdf";

            // Save the workbook using the configured options (only the selected sheet will be rendered)
            workbook.Save(pdfFileName, pdfOptions);
        }
    }
}