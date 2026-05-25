using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create

        // Example of loading an existing workbook:
        // Workbook workbook = new Workbook("input.xlsx"); // load

        // Add sample data and a formula (only needed if the workbook was created)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula(); // required step before PDF export

        // Set PDF save options (formula calculation already done, so keep false)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CalculateFormula = false
        };

        // Save the workbook as PDF
        using (MemoryStream pdfStream = new MemoryStream())
        {
            workbook.Save(pdfStream, pdfOptions); // save
            File.WriteAllBytes("Result.pdf", pdfStream.ToArray());
        }
    }
}