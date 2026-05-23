using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfFormExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file that contains form fields
            string sourcePath = "input.xlsx";

            // Path for the generated PDF file
            string pdfPath = "output.pdf";

            // Load the workbook from the Excel file (create/load lifecycle rule)
            Workbook workbook = new Workbook(sourcePath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Enable form‑field filling permission so that the PDF retains interactive fields
            PdfSecurityOptions security = new PdfSecurityOptions
            {
                FillFormsPermission = true   // Allows users to fill existing form fields
            };
            pdfOptions.SecurityOptions = security;

            // Save the workbook as PDF while preserving the interactive form fields
            // (save lifecycle rule using Save(string, SaveOptions))
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"Workbook converted to PDF with interactive form fields: {pdfPath}");
        }
    }
}