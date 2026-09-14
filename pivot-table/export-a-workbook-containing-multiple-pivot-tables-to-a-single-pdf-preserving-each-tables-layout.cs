// Title: Export a workbook with multiple pivot tables to a single PDF while preserving each table’s layout using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that builds a workbook, adds a tabular‑form pivot table and an outline‑form pivot table on separate sheets, and saves the whole workbook as one PDF with each pivot table’s layout retained. | Show how to configure Aspose.Cells PdfSaveOptions to export a workbook containing several pivot tables to a single PDF, ensuring the document structure and pivot formatting are preserved.
// Common Searches: C# Aspose.Cells export workbook with multiple pivot tables to one PDF file | how to keep pivot table formatting when saving to PDF using Aspose.Cells | Aspose.Cells PdfSaveOptions ExportDocumentStructure example for pivot tables | save multiple worksheets as a single PDF while preserving pivot layouts in .NET | export tabular and outline pivot tables to PDF with Aspose.Cells
// Tags: Aspose.Cells export workbook to single PDF | preserve pivot table layout in PDF export | PdfSaveOptions ExportDocumentStructure true | create tabular pivot table Aspose.Cells | create outline pivot table Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook with sample data, adds two pivot tables—one displayed in tabular form and another in outline form—on separate worksheets, refreshes the pivots, configures PdfSaveOptions to retain document structure, and saves the entire workbook as a single PDF that keeps each pivot table’s layout intact.
class ExportPivotTablesToPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // 1. Prepare source data for the pivot tables
        // -------------------------------------------------
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Header row
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Product");
        dataSheet.Cells["C1"].PutValue("Sales");

        // Sample rows
        string[] categories = { "Fruit", "Fruit", "Vegetable", "Vegetable" };
        string[] products   = { "Apple", "Banana", "Carrot", "Potato" };
        int[]    sales      = { 1200, 800, 600, 400 };

        for (int i = 0; i < categories.Length; i++)
        {
            dataSheet.Cells[i + 1, 0].PutValue(categories[i]); // Column A
            dataSheet.Cells[i + 1, 1].PutValue(products[i]);   // Column B
            dataSheet.Cells[i + 1, 2].PutValue(sales[i]);     // Column C
        }

        // -------------------------------------------------
        // 2. Create first pivot table (tabular form)
        // -------------------------------------------------
        Worksheet pivotSheet1 = workbook.Worksheets.Add("PivotTable1");
        int ptIndex1 = pivotSheet1.PivotTables.Add("=Data!A1:C5", "A3", "Pivot1");
        PivotTable pivot1 = pivotSheet1.PivotTables[ptIndex1];
        pivot1.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot1.AddFieldToArea(PivotFieldType.Column, "Product");
        pivot1.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot1.ShowInTabularForm(); // Preserve tabular layout

        // -------------------------------------------------
        // 3. Create second pivot table (outline form)
        // -------------------------------------------------
        Worksheet pivotSheet2 = workbook.Worksheets.Add("PivotTable2");
        int ptIndex2 = pivotSheet2.PivotTables.Add("=Data!A1:C5", "A3", "Pivot2");
        PivotTable pivot2 = pivotSheet2.PivotTables[ptIndex2];
        pivot2.AddFieldToArea(PivotFieldType.Row, "Product");
        pivot2.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot2.ShowInOutlineForm(); // Preserve outline layout

        // -------------------------------------------------
        // 4. Refresh all pivot tables to ensure data is up‑to‑date
        // -------------------------------------------------
        workbook.Worksheets.RefreshPivotTables();

        // -------------------------------------------------
        // 5. Configure PDF save options
        // -------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Keep each worksheet on its own pages (default behavior)
            OnePagePerSheet = false,
            // Export the document structure so that each pivot table retains its layout
            ExportDocumentStructure = true
        };

        // -------------------------------------------------
        // 6. Save the workbook as a single PDF file
        // -------------------------------------------------
        workbook.Save("MultiplePivotTables.pdf", pdfOptions);
    }
}
