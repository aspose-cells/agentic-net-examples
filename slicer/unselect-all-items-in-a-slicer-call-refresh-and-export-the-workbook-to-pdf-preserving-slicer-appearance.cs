// Title: C# – Deselect All Items in an Aspose.Cells Slicer, Refresh the Pivot, and Export Workbook to PDF with Slicer Visuals
// Description: Shows how to clear slicer filters with ShowAllItems, refresh the linked pivot table, and save the workbook as a PDF using Aspose.Cells PdfSaveOptions so the slicer graphic is retained.
// Keywords: Aspose.Cells slicer clear selection | ShowAllItems Aspose.Cells | Refresh slicer C# | PivotTable refresh Aspose.Cells | Export workbook to PDF Aspose.Cells | PdfSaveOptions slicer appearance | C# Aspose.Cells PDF export | Unselect all slicer items | Preserve slicer visual in PDF
// Common Searches: how to clear slicer selection Aspose.Cells C# | Aspose.Cells refresh slicer before PDF export | export workbook to PDF keeping slicer graphic | C# Aspose.Cells ShowAllItems example | save Excel workbook with slicer as PDF
// Developer Intent: The developer needs to programmatically remove any slicer filters, refresh the associated pivot table, and generate a PDF that still displays the slicer control.
// Use Cases: Create a sales dashboard PDF that always shows all categories after resetting slicer filters. | Automate monthly report generation where the slicer visual must appear unchanged in the exported PDF. | Provide an unfiltered view of pivot data in PDF form for client presentations.
// AI Prompts: Generate C# code with Aspose.Cells to deselect all items in a slicer, refresh the linked pivot table, and export the workbook to PDF preserving the slicer layout. | Explain the role of ShowAllItems and Refresh when preparing a workbook for PDF export with Aspose.Cells. | Show how to configure PdfSaveOptions so slicer graphics are included in the PDF output.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

// Shows how to clear slicer filters with ShowAllItems, refresh the linked pivot table, and save the workbook as a PDF using Aspose.Cells PdfSaveOptions so the slicer graphic is retained.
class ExportWorkbookWithSlicer
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].PutValue("Fruit");
        cells["B1"].PutValue("Sales");
        cells["A2"].PutValue("Apple");
        cells["B2"].PutValue(120);
        cells["A3"].PutValue("Banana");
        cells["B3"].PutValue(80);
        cells["A4"].PutValue("Orange");
        cells["B4"].PutValue(150);
        cells["A5"].PutValue("Apple");
        cells["B5"].PutValue(90);

        // Add a pivot table based on the data range
        int pivotIdx = sheet.PivotTables.Add("A1:B5", "D2", "FruitPivot");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the pivot table for the "Fruit" field
        SlicerCollection slicers = sheet.Slicers;
        int slicerIdx = slicers.Add(pivot, "F2", "Fruit");
        Slicer slicer = slicers[slicerIdx];

        // Unselect all items (show all items) and refresh the slicer
        slicer.ShowAllItems = true;   // clears any filter applied by the slicer
        slicer.Refresh();             // refreshes slicer and underlying pivot table

        // Prepare PDF save options to preserve slicer appearance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure all visible sheets are exported (default)
            SheetSet = SheetSet.Visible
        };

        // Export the workbook to PDF
        workbook.Save("WorkbookWithSlicer.pdf", pdfOptions);
    }
}
