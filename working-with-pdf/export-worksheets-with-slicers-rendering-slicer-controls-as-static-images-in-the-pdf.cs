// Title: Export Worksheet with Pivot Slicer to PDF as Static Image – Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, builds a pivot table, inserts a slicer, marks it printable, and saves the first sheet to PDF using PdfSaveOptions (OnePagePerSheet, SheetSet). The slicer is rendered as a static image in the resulting PDF.
// Keywords: Aspose.Cells export slicer PDF | render slicer as image PDF | PdfSaveOptions OnePagePerSheet | slicer printable property | pivot table PDF export .NET
// Common Searches: Aspose.Cells export slicer to PDF | how to make slicer appear in PDF with Aspose.Cells | PdfSaveOptions sheetset slicer image | render Excel slicer as static image in PDF | C# Aspose.Cells slicer printable PDF
// Developer Intent: Save an Excel workbook that contains a slicer so the slicer is displayed as a static image in the generated PDF file.
// Use Cases: Produce a PDF snapshot of a dashboard sheet that includes pivot tables and slicer controls for offline review. | Distribute a one‑page PDF report showing slicer selections alongside data visualizations. | Export only selected worksheets that contain slicers while preserving their layout and appearance.
// AI Prompts: Generate C# code with Aspose.Cells to export a worksheet containing a slicer to PDF, ensuring the slicer is rendered as a static image. | Explain the effect of the Slicer.IsPrintable property on PDF rendering in Aspose.Cells and suggest current best‑practice alternatives. | Show how to configure PdfSaveOptions to export multiple sheets with slicers, each on a separate PDF page.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

namespace AsposeCellsSlicerPdfExport
{
    // Creates a workbook, adds sample data, builds a pivot table, inserts a slicer, marks it printable, and saves the first sheet to PDF using PdfSaveOptions (OnePagePerSheet, SheetSet). The slicer is rendered as a static image in the resulting PDF.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample data
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["B1"].PutValue("Year");
            sheet.Cells["C1"].PutValue("Amount");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(2020);
            sheet.Cells["C2"].PutValue(50);

            sheet.Cells["A3"].PutValue("Apple");
            sheet.Cells["B3"].PutValue(2021);
            sheet.Cells["C3"].PutValue(70);

            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(2020);
            sheet.Cells["C4"].PutValue(30);

            sheet.Cells["A5"].PutValue("Banana");
            sheet.Cells["B5"].PutValue(2021);
            sheet.Cells["C5"].PutValue(60);

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:C5", "E2", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "Year");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table (field: Fruit)
            int slicerIdx = sheet.Slicers.Add(pivot, "G2", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Ensure the slicer is printable so it appears in the PDF as a static image
            slicer.IsPrintable = true; // obsolete but still functional for rendering

            // ---------- Configure PDF save options ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,               // render each sheet on a single PDF page
                // Optional: render only the first sheet
                SheetSet = new SheetSet(new int[] { 0 })
            };

            // ---------- Save the workbook as PDF ----------
            workbook.Save("WorkbookWithSlicer.pdf", pdfOptions);
        }
    }
}
