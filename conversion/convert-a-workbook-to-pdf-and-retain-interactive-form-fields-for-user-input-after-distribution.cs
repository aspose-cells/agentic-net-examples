// Title: Convert Excel to PDF with Fillable Form Fields using Aspose.Cells for .NET
// Description: This example loads an Excel workbook, configures PdfSaveOptions with PdfSecurityOptions (FillFormsPermission and PrintPermission), enables ExportDocumentStructure, and saves the file as a PDF. The resulting PDF retains the original interactive form fields, allowing users to fill them after distribution.
// Keywords: Aspose.Cells PDF conversion | fillable PDF from Excel | PdfSaveOptions FillFormsPermission | ExportDocumentStructure | PDF security options .NET | preserve Excel form controls | Aspose.Cells C# example
// Common Searches: Aspose.Cells keep Excel form controls when saving as PDF | How to enable fillable fields in PDF generated from Excel | PdfSecurityOptions FillFormsPermission C# | ExportDocumentStructure preserve interactive elements Aspose.Cells | Convert Excel to PDF with printable and fillable forms
// Developer Intent: Convert an Excel workbook to a PDF while retaining its fillable form fields for end‑user input.
// Use Cases: Create fillable PDF contracts from Excel templates for client distribution. | Generate PDF questionnaires that users can complete electronically. | Produce PDF reports with embedded data‑entry fields while restricting other edits. | Automate batch conversion of Excel forms to fillable PDFs for onboarding workflows.
// AI Prompts: Show how to add a password to the PDF while preserving FillFormsPermission. | Provide code to merge multiple worksheets into a single fillable PDF. | Explain how to disable content editing but keep form filling enabled. | Demonstrate setting PDF permissions for printing, copying, and form filling in Aspose.Cells. | Give an example of converting a workbook with charts and preserving interactive fields.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfFormExample
{
    // This example loads an Excel workbook, configures PdfSaveOptions with PdfSecurityOptions (FillFormsPermission and PrintPermission), enables ExportDocumentStructure, and saves the file as a PDF. The resulting PDF retains the original interactive form fields, allowing users to fill them after distribution.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file that contains interactive form fields
            string sourcePath = "input.xlsx";

            // Path where the resulting PDF will be saved
            string destPath = "output.pdf";

            // Load the workbook (uses the provided create/load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Configure PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Set security options to allow filling form fields after distribution
            PdfSecurityOptions securityOptions = new PdfSecurityOptions
            {
                // Enable form filling permission even if other modify permissions are disabled
                FillFormsPermission = true,

                // Optional: allow printing of the PDF
                PrintPermission = true
            };

            // Assign the security options to the PDF save options
            pdfSaveOptions.SecurityOptions = securityOptions;

            // Optional: retain the document structure (helps preserve interactive elements)
            pdfSaveOptions.ExportDocumentStructure = true;

            // Save the workbook as PDF using the provided save rule with SaveOptions
            workbook.Save(destPath, pdfSaveOptions);

            Console.WriteLine($"Workbook converted to PDF with interactive form fields retained: {destPath}");
        }
    }
}
