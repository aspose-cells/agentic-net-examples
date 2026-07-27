using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data with a formula
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula(new CalculationOptions());

        // Configure PDF save options to ensure formulas are calculated during save
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            CalculateFormula = true
        };

        // Save the workbook as PDF to a memory stream, then write to a file
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, pdfSaveOptions);
            File.WriteAllBytes("output.pdf", stream.ToArray());
        }
    }
}

// Author: Aspose.Cells example code demonstrating formula calculation before PDF export.