// Title: Validate Workbook Title Built‑In Property Before Export with Aspose.Cells for .NET
// Description: Demonstrates how to create a Workbook, assign a value to the built‑in Title property, verify that the title is not null, empty, or whitespace, raise an InvalidOperationException if it is, and then export the workbook to PDF using PdfSaveOptions with DisplayDocTitle enabled.
// Keywords: Aspose.Cells .NET | built‑in document properties | title validation | export workbook to PDF | DisplayDocTitle | metadata compliance | prevent empty title | save workbook with title check | document title PDF viewer
// Common Searches: Aspose.Cells check Title property before saving | how to enforce workbook title in .NET | export to PDF only when document title exists | throw error if workbook title is missing Aspose.Cells | display document title in PDF window Aspose.Cells
// Developer Intent: Ensure the workbook's Title built‑in property is set and non‑empty before any export operation.
// Use Cases: Guarantee that generated reports contain a title for regulatory or archival purposes. | Show the report name in the PDF viewer's title bar by enabling DisplayDocTitle after validation. | Prevent accidental creation of untitled files in automated document pipelines.
// AI Prompts: Write a C# method that checks a Workbook's Title property and throws an exception if it is blank before calling Save. | Provide Aspose.Cells code that exports a workbook to PDF only after confirming the built‑in Title is populated, with DisplayDocTitle set to true. | Create unit tests that verify the title‑validation logic blocks saving when the Title property is missing or whitespace.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace Example
{
    // Demonstrates how to create a Workbook, assign a value to the built‑in Title property, verify that the title is not null, empty, or whitespace, raise an InvalidOperationException if it is, and then export the workbook to PDF using PdfSaveOptions with DisplayDocTitle enabled.
    class ValidateTitleBeforeExport
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create rule)
                Workbook workbook = new Workbook();

                // Set the built‑in Title property (required before export)
                workbook.BuiltInDocumentProperties.Title = "Quarterly Report";

                // Validate that the Title property is not empty
                string title = workbook.BuiltInDocumentProperties.Title;
                if (string.IsNullOrWhiteSpace(title))
                {
                    throw new InvalidOperationException("The workbook's Title property must be set before exporting.");
                }

                // Export the workbook to PDF (lifecycle save rule)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Display the document title in the PDF window title bar
                    DisplayDocTitle = true
                };
                workbook.Save("ExportedReport.pdf", pdfOptions);
                Console.WriteLine("Workbook exported successfully to ExportedReport.pdf");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
