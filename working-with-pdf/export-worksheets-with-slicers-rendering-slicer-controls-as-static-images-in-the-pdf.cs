// Title: Export Pivot Table Slicer as a Static Image to PDF with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, builds a pivot table, inserts a linked slicer, marks the slicer shape as printable, and saves the file to PDF using PdfSaveOptions so the slicer appears as a static image on the PDF page.
// Keywords: Aspose.Cells | C# | PDF export | slicer | pivot table | static image | IsPrintable | PdfSaveOptions | export worksheet to PDF
// Common Searches: Aspose.Cells export slicer to PDF | make slicer printable in PDF using Aspose.Cells | C# save workbook with pivot slicer as image PDF | render slicer as image in PDF Aspose.Cells | one page per sheet PDF export Aspose.Cells
// Developer Intent: Generate a PDF where pivot table slicer controls are rendered as non‑interactive images.
// Use Cases: Produce printable PDF reports that include both pivot tables and their slicer selections. | Automate dashboard snapshots where slicer states are captured as static graphics for archival. | Export multiple worksheets, each containing slicers, to a single PDF with one page per sheet.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook containing a pivot table slicer to PDF, ensuring the slicer is rendered as a static image. | Explain how to set the slicer shape's IsPrintable property and configure PdfSaveOptions for one page per sheet during PDF export. | Provide a C# loop that iterates through all worksheets in a workbook and saves them together in a combined PDF while preserving slicer images.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering; // For PdfSaveOptions if needed

namespace ExportSlicersToPdf
{
    // Creates a workbook, adds sample data, builds a pivot table, inserts a linked slicer, marks the slicer shape as printable, and saves the file to PDF using PdfSaveOptions so the slicer appears as a static image on the PDF page.
    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook --------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(20);

            // -------------------- Create pivot table --------------------
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D2", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------- Add slicer linked to the pivot --------------------
            // Destination cell for slicer upper‑left corner
            string slicerCell = "F2";
            // Base field name must exist in the pivot's BaseFields collection
            int slicerIdx = sheet.Slicers.Add(pivot, slicerCell, "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Ensure the slicer is printable so it appears in the PDF as a static image
            // The modern way is to use the Shape object's IsPrintable property
            slicer.Shape.IsPrintable = true;

            // -------------------- Export workbook to PDF --------------------
            // PdfSaveOptions can be used to control rendering; default settings already render shapes.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: export each worksheet on a separate page
                OnePagePerSheet = true
            };

            // Save the workbook as PDF; slicer will be rendered as a static image.
            workbook.Save("WorkbookWithSlicers.pdf", pdfOptions);
        }
    }
}
