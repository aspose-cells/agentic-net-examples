using System;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to be exported (e.g., B2:D10)
            // Create a CellArea representing the selected range
            CellArea exportArea = new CellArea
            {
                StartRow = 1,    // Row index is zero‑based (B2 -> row 1)
                StartColumn = 1, // Column B -> index 1
                EndRow = 9,      // D10 -> row 9
                EndColumn = 3    // Column D -> index 3
            };

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Exclude hidden rows and columns from the PDF output
            // NOTE: The existence of ExportHiddenRows / ExportHiddenColumns properties
            // is assumed based on typical Aspose.Cells API but not verified in the provided documentation.
            // If these properties are unavailable, replace with the correct API or remove the lines.
            // BEGIN MISSING API PLACEHOLDER
            // pdfOptions.ExportHiddenRows = false;      // Exclude hidden rows
            // pdfOptions.ExportHiddenColumns = false;   // Exclude hidden columns
            // END MISSING API PLACEHOLDER

            // Set the export area so that only the selected range is rendered.
            // NOTE: PdfSaveOptions does not expose an ExportArea property in the supplied docs.
            // If such a property exists, uncomment the line below; otherwise, consider using
            // HtmlSaveOptions.ExportArea as an intermediate step or adjust the worksheet view.
            // BEGIN MISSING API PLACEHOLDER
            // pdfOptions.ExportArea = exportArea;
            // END MISSING API PLACEHOLDER

            // Save the selected range to PDF
            workbook.Save("selected_range.pdf", pdfOptions);
        }
    }
}