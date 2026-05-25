using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions if needed

class Program
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some data
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        // Add a formula that sums A1 and A2
        cells["A3"].Formula = "=SUM(A1:A2)";

        // Explicitly calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Prepare PDF save options (optional: CalculateFormula can be left false because we already calculated)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CalculateFormula = false // formulas already evaluated
        };

        // Save the workbook as PDF
        string outputPath = "CalculatedWorkbook.pdf";
        workbook.Save(outputPath, pdfOptions);

        Console.WriteLine($"Workbook saved to '{outputPath}'. Formula result in A3: {cells["A3"].Value}");
    }
}