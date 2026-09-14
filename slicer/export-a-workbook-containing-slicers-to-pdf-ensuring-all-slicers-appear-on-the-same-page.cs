// Title: Export an Aspose.Cells workbook with slicers to a single‑page PDF in C#
// AI Prompts: Generate C# code that creates a worksheet, adds a table and a slicer, marks the slicer printable, and saves the workbook as a PDF using PdfSaveOptions so everything fits on one page. | Show how to configure Aspose.Cells PdfSaveOptions with OnePagePerSheet and AllColumnsInOnePagePerSheet to force all worksheet content, including slicers, onto a single PDF page. | Explain the steps required to make a slicer printable in Aspose.Cells before exporting the workbook to PDF.
// Common Searches: Aspose.Cells C# export workbook with slicer to one-page PDF | How to keep slicer visible when saving Excel to PDF using Aspose.Cells | PdfSaveOptions OnePagePerSheet slicer printable Aspose.Cells example | Fit all columns and slicers on a single PDF page with Aspose.Cells | C# code sample for adding slicer and exporting to PDF in Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | Aspose.Cells slicer printable shape | export slicer to PDF Aspose.Cells C# | fit all columns single PDF page Aspose.Cells | create table and slicer Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

namespace ExportWorkbookWithSlicers
{
    // The example creates a new workbook, adds sample data and a table, inserts a slicer linked to the first column, sets the slicer’s Shape.IsPrintable property, configures PdfSaveOptions with OnePagePerSheet and AllColumnsInOnePagePerSheet, and saves the workbook as a PDF where the slicer appears on the same page as the data.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data for a table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("C");
                sheet.Cells["B5"].PutValue(40);

                // Add a table covering the data range
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.ShowHeaderRow = true;
                // Total row not required for this example; omitted to avoid API mismatch

                // Add a slicer linked to the first column of the table
                // Position the slicer at D2 (row 1, column 3)
                int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[0], 1, 3);
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Ensure the slicer is printable (required for PDF export)
                slicer.Shape.IsPrintable = true;

                // Configure PDF save options to force all content of the sheet onto a single page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true,               // All content on one page
                    AllColumnsInOnePagePerSheet = true    // Fit all columns within the page width
                };

                // Save the workbook as PDF; slicer will appear on the same page as the data
                workbook.Save("WorkbookWithSlicers.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
