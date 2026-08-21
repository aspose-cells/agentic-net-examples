// Title: Export Aspose.Cells Workbook with Slicers to a Single‑Page PDF (C#)
// Description: Creates a workbook, adds a table and a slicer, marks the slicer as printable, and saves the file as PDF using PdfSaveOptions with OnePagePerSheet and AllColumnsInOnePagePerSheet so the slicer and data appear on the same page.
// Keywords: Aspose.Cells | C# | slicer PDF export | PdfSaveOptions | OnePagePerSheet | AllColumnsInOnePagePerSheet | printable slicer | single page PDF | Excel slicer to PDF | Aspose.Cells PDF export
// Common Searches: Aspose.Cells export slicer to PDF | PDFSaveOptions single page sheet | make slicer printable Aspose.Cells | C# export workbook with slicer as PDF | force all content onto one PDF page Aspose
// Developer Intent: Generate a PDF from an Aspose.Cells workbook that contains slicers, ensuring the slicers are rendered on the same page as the worksheet data.
// Use Cases: Produce a printable dashboard where a table and its slicer are shown together on one PDF page. | Automate conversion of Excel reports with slicers into single‑page PDFs for distribution or archiving. | Create a PDF snapshot of an interactive Excel slicer layout for compliance documentation.
// AI Prompts: How do I export an Aspose.Cells workbook with slicers to a single‑page PDF in C#? | Provide C# code that sets slicer shapes to printable and uses PdfSaveOptions to keep all content on one PDF page. | Explain the effect of OnePagePerSheet and AllColumnsInOnePagePerSheet on slicer rendering during PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;
using PdfSaveOptions = Aspose.Cells.PdfSaveOptions;

namespace AsposeCellsSlicerPdfExport
{
    // Creates a workbook, adds a table and a slicer, marks the slicer as printable, and saves the file as PDF using PdfSaveOptions with OnePagePerSheet and AllColumnsInOnePagePerSheet so the slicer and data appear on the same page.
    public class ExportSlicersToPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("A");
                worksheet.Cells["B4"].PutValue(30);
                worksheet.Cells["A5"].PutValue("C");
                worksheet.Cells["B5"].PutValue(40);

                // Add a table that covers the data range (A1:B5)
                int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                // Set display name for the table (Name property may not be available in this version)
                table.DisplayName = "SampleTable";

                // Add a slicer for the "Category" column of the table
                // Position the slicer at cell D2 (row 1, column 3)
                int slicerIndex = worksheet.Slicers.Add(table, table.ListColumns[0], 1, 3);
                Slicer slicer = worksheet.Slicers[slicerIndex];
                // Ensure slicer appears in PDF (use Shape.IsPrintable)
                slicer.Shape.IsPrintable = true;

                // Configure PDF save options to force all content onto a single page per sheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true,
                    AllColumnsInOnePagePerSheet = true
                };

                // Save the workbook as PDF; slicers will appear on the same page as the worksheet data
                workbook.Save("SlicersExported.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                // Log or handle exceptions as needed
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportSlicersToPdf.Run();
        }
    }
}
