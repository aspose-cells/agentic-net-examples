// Title: Hide columns 5‑9 in an Aspose.Cells worksheet and export to PDF while keeping them hidden (C#)
// Common Searches: Aspose.Cells hide columns 5 to 9 C# | export worksheet to PDF while preserving hidden columns using Aspose.Cells | how to hide a range of columns in Aspose.Cells before PDF conversion | C# hide multiple columns with Cells.HideColumn in Aspose.Cells | keep hidden column state when saving workbook as PDF with Aspose.Cells
// Tags: hide column range Aspose.Cells | pdf export hidden columns Aspose.Cells | PdfSaveOptions column visibility | C# hide multiple columns zero based index | preserve hidden column state Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example creates a workbook, fills it with sample data, hides columns indexed 5 through 9 using Cells.HideColumn, configures PdfSaveOptions, and saves the worksheet as a PDF while retaining the hidden column state.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                for (int row = 0; row < 20; row++)
                {
                    for (int col = 0; col < 15; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Hide columns with indexes 5 to 9 (zero‑based indexing)
                for (int colIndex = 5; colIndex <= 9; colIndex++)
                {
                    sheet.Cells.HideColumn(colIndex);
                }

                // Prepare PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // If the current Aspose.Cells version supports this property, enable it:
                // pdfOptions.IsHiddenColumnVisible = true;

                // Export the workbook to PDF
                string outputPath = "ExportedWithHiddenColumns.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
