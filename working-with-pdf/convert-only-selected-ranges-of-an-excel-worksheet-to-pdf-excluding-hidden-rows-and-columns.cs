using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    class ConvertSelectedRangeToPdf
    {
        static void Main()
        {
            // Load the source Excel workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can change the index as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example: hide some rows and columns that should be excluded from the PDF
            // Hide row 3 (zero‑based index)
            cells.HideRow(2);
            // Hide column B (zero‑based index)
            cells.HideColumn(1);

            // Define the range that you want to convert to PDF.
            // This range will be set as the print area, so only this area is exported.
            // Adjust the address as required (e.g., "A1:D20").
            sheet.PageSetup.PrintArea = "A1:D20";

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Optional: ignore completely blank pages in the output PDF
            pdfOptions.PrintingPageType = PrintingPageType.IgnoreBlank;

            // Save the selected range to PDF.
            // Hidden rows and columns within the defined print area are automatically omitted.
            workbook.Save("selected_range_output.pdf", pdfOptions);
        }
    }
}