// Title: Convert an Aspose.Cells workbook to PDF with fillable form fields using C#
// AI Prompts: Generate C# code that adds text box shapes to a worksheet and saves the workbook as a PDF with interactive fillable fields using Aspose.Cells. | Show how to configure PdfSaveOptions and PdfSecurityOptions in Aspose.Cells to enable FillFormsPermission and export document structure for accessibility. | Provide a C# example that creates the output directory if it does not exist and writes the PDF with accessible, editable form fields.
// Common Searches: how to keep Excel text boxes as fillable fields when exporting to PDF with Aspose.Cells C# | Aspose.Cells PDF export enable fill forms permission C# example | preserve interactive form fields in PDF generated from workbook using Aspose.Cells | C# Aspose.Cells create accessible PDF with document structure and fillable forms | save workbook as PDF with editable form fields Aspose.Cells .NET
// Tags: Aspose.Cells PDF export with fillable form fields | PdfSecurityOptions FillFormsPermission C# | PdfSaveOptions export document structure Aspose.Cells | add text box shape as PDF form field Aspose.Cells | C# create accessible PDF from Excel workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;               // For Shape class
using Aspose.Cells.Rendering.PdfSecurity; // For PDF security options

// The program creates a new workbook, inserts named text box shapes for Name and Email, configures PdfSaveOptions with FillFormsPermission and document structure export, ensures the output directory exists, and saves the workbook as 'InteractiveForm.pdf', producing a PDF that retains interactive, fillable form fields for end users.
class WorkbookToPdfWithInteractiveForms
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Name:");
            sheet.Cells["A2"].PutValue("Email:");

            // Add text box shapes that will become interactive PDF fields
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, height, width
            Shape nameShape = sheet.Shapes.AddTextBox(0, 1, 0, 0, 30, 150);
            nameShape.Name = "NameField";

            Shape emailShape = sheet.Shapes.AddTextBox(1, 1, 0, 0, 30, 150);
            emailShape.Name = "EmailField";

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Enable form filling permission so users can fill the fields after distribution
            PdfSecurityOptions security = new PdfSecurityOptions
            {
                FillFormsPermission = true // Allow filling existing interactive form fields
            };
            pdfOptions.SecurityOptions = security;

            // Preserve document structure for better accessibility (optional)
            pdfOptions.ExportDocumentStructure = true;

            // Define output path and ensure the directory exists
            string outputPath = "InteractiveForm.pdf";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF with interactive form fields: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
