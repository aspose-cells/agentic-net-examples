using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets[0].Cells["A1"].PutValue("Data in Sheet1");

        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].PutValue("Data in Sheet2");

        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
        sheet3.Cells["A1"].PutValue("Data in Sheet3");

        // Loop through each worksheet, set SheetSet to the current sheet,
        // and save it as an individual PDF file.
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            // Configure PDF save options for a single sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.SheetSet = new SheetSet(new int[] { i }); // select only the i‑th sheet

            // Generate a file name like "Sheet_1.pdf", "Sheet_2.pdf", etc.
            string outputFile = $"Sheet_{i + 1}.pdf";

            // Save the workbook (only the selected sheet) to PDF
            workbook.Save(outputFile, pdfOptions);
        }
    }
}