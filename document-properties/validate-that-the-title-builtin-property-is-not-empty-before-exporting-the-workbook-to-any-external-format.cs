// Title: Validate Workbook Title Built‑In Property Before Export with Aspere.Cells for .NET
// Description: Demonstrates how to set the built‑in Title property of an Aspose.Cells Workbook, verify that it is not null, empty, or whitespace, and abort the export if the check fails. The example also shows configuring PdfSaveOptions.DisplayDocTitle for PDF output, but the validation logic can be applied to any export format.
// Keywords: Aspose.Cells title validation | C# workbook built‑in document properties | check Title before export | PdfSaveOptions DisplayDocTitle requirement | export Excel to PDF with title check | .NET document property validation | prevent empty Title in Aspose export
// Common Searches: how to ensure Title property is set before saving workbook as PDF using Aspose.Cells | Aspose.Cells .NET throw error when document Title is missing | DisplayDocTitle option needs non‑empty Title Aspose.Cells | validate built‑in document properties before exporting Excel file | C# Aspose.Cells check Title property before any format export
// Developer Intent: Confirm that the workbook’s Title built‑in property contains a value before performing any export operation.
// Use Cases: Enforce compliance by blocking PDF generation when the Title metadata is absent. | Add a pre‑export validation step in a batch process that converts many Excel files to PDFs or other formats. | Integrate title verification into a web service that returns generated documents, returning a clear error message if the Title is missing.
// AI Prompts: Create a reusable C# method that validates required built‑in document properties (Title, Author, etc.) before saving a workbook with Aspose.Cells. | Generate code that logs a missing Title property and skips the export in a loop processing multiple workbooks. | Provide an example that sets PdfSaveOptions.DisplayDocTitle only after confirming the workbook Title is not empty.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to set the built‑in Title property of an Aspose.Cells Workbook, verify that it is not null, empty, or whitespace, and abort the export if the check fails. The example also shows configuring PdfSaveOptions.DisplayDocTitle for PDF output, but the validation logic can be applied to any export format.
class ValidateTitleBeforeExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Set the Title property (required for PDF export with DisplayDocTitle)
            workbook.BuiltInDocumentProperties.Title = "Sample Document Title";

            // Validate that the built‑in Title property is not empty or whitespace
            if (string.IsNullOrWhiteSpace(workbook.BuiltInDocumentProperties.Title))
            {
                throw new InvalidOperationException("The workbook Title property must be set before exporting.");
            }

            // Prepare save options for PDF format
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the document title appears in the PDF window title bar
                DisplayDocTitle = true
            };

            // Export the workbook using the validated Title
            workbook.Save("ExportedDocument.pdf", pdfOptions);
            Console.WriteLine("Workbook exported successfully to ExportedDocument.pdf");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
