using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // Configure PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            // Retain document structure in the generated PDF
            ExportDocumentStructure = true
            // Additional options can be set here as needed
        };

        // Ensure formulas are calculated before saving (optional)
        workbook.CalculateFormula();

        // Save the workbook to PDF using the configured options
        string outputPath = "output.pdf";
        workbook.Save(outputPath, pdfSaveOptions);
    }
}

// Author: Aspose.Cells .NET example code.