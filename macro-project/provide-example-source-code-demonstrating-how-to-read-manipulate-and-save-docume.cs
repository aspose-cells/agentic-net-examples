using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Manipulate the workbook: add a new sheet and write some data
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet = workbook.Worksheets[sheetIndex];
            sheet.Name = "Report";
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(85);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set PDF/A-1b compliance
            pdfOptions.Compliance = PdfCompliance.PdfA1b;

            // Export document structure (useful for accessibility)
            pdfOptions.ExportDocumentStructure = true;

            // Set a custom producer string
            pdfOptions.Producer = "Aspose.Cells PDF Demo";

            // Configure security options (password protection, permissions)
            PdfSecurityOptions security = new PdfSecurityOptions
            {
                OwnerPassword = "ownerPass",
                UserPassword = "userPass",
                PrintPermission = true,
                ModifyDocumentPermission = false,
                ExtractContentPermission = false,
                AnnotationsPermission = true,
                FillFormsPermission = true
            };
            pdfOptions.SecurityOptions = security;

            // Save the manipulated workbook as a PDF with the specified options
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("Workbook has been read, modified, and saved as PDF.");
        }
    }
}