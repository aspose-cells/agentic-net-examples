// Title: C# – Remove All Slicers from an Aspose.Cells Workbook and Export to PDF
// Description: A complete C# example that loads or creates an Excel workbook, iterates through every worksheet, clears all slicers with `Worksheet.Slicers.Clear()`, configures `PdfSaveOptions` to include all visible sheets, and saves the result as a PDF file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | remove slicer | clear slicers | delete slicers | iterate worksheets | PdfSaveOptions | export to PDF | Workbook.Save | Slicers.Clear | Aspose.Cells example | GitHub sample | code snippet
// Common Searches: how to delete all slicers in Aspose.Cells C# | Aspose.Cells remove slicers before PDF export | clear slicers from every worksheet programmatically | C# Aspose.Cells export workbook to PDF after removing slicers | PdfSaveOptions SheetSet.All example
// Developer Intent: Programmatically clear every slicer in a workbook and generate a PDF version.
// Use Cases: Prepare a clean printable PDF of a dashboard by stripping interactive slicers. | Automate batch processing of reports where slicers are not needed in the final PDF. | Integrate slicer removal into a CI pipeline that generates documentation PDFs from Excel templates.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, removes all slicers from each worksheet, and saves the workbook as a PDF. | Explain how to use PdfSaveOptions.SheetSet = SheetSet.All to export every sheet after clearing slicers. | Show error‑handling best practices when deleting slicers and exporting to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;
using Aspose.Cells.Tables;   // Required for ListObject

namespace SlicerRemovalAndPdfExport
{
    // A complete C# example that loads or creates an Excel workbook, iterates through every worksheet, clears all slicers with `Worksheet.Slicers.Clear()`, configures `PdfSaveOptions` to include all visible sheets, and saves the result as a PDF file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load an existing file

                // Example data and slicer creation (optional, for demonstration)
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(20);

                // Add a table (ListObject) covering the data range
                int tableIdx = ws.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = ws.ListObjects[tableIdx];

                // Add a slicer linked to the first column of the table
                ws.Slicers.Add(table, table.ListColumns[0], 1, 3);

                // Iterate all worksheets and delete every slicer
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Clear removes all slicers from the worksheet
                    sheet.Slicers.Clear();
                }

                // Prepare PDF save options (optional customizations)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Export all visible sheets; default is SheetSet.Visible
                    SheetSet = SheetSet.All
                };

                // Export the modified workbook to PDF
                string outputPdf = "WorkbookWithoutSlicers.pdf";
                workbook.Save(outputPdf, pdfOptions);
                Console.WriteLine($"Workbook saved as PDF: {Path.GetFullPath(outputPdf)}");
            }
            catch (Exception ex)
            {
                // Log or display the error details
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
