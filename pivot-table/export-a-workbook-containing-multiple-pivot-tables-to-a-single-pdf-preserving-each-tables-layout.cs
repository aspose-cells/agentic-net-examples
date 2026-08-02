// Title: Export Multiple Pivot Tables to One PDF with Layout Preservation – Aspose.Cells for .NET (C#)
// Description: Creates a workbook with two data sheets, adds a pivot table on each sheet, refreshes them, configures PdfSaveOptions (ExportDocumentStructure = true, OnePagePerSheet = false) and saves all sheets as a single PDF while keeping each pivot's layout intact.
// Keywords: Aspose.Cells | C# | pivot table PDF export | multiple pivot tables | ExportDocumentStructure | OnePagePerSheet | PdfSaveOptions | save workbook as PDF | preserve pivot layout | refresh pivot tables | Aspose.Cells .NET example
// Common Searches: Aspose.Cells export multiple pivot tables to PDF | preserve pivot table layout when saving as PDF C# | PdfSaveOptions ExportDocumentStructure example | save entire workbook as single PDF Aspose.Cells | C# code to export pivot tables to PDF
// Developer Intent: Produce a single PDF that contains all pivot tables from a workbook without losing their formatting or layout.
// Use Cases: Generate a consolidated financial report where each department’s pivot summary appears on its own sheet in one PDF. | Create a sales dashboard PDF that combines product‑by‑region pivot tables from separate worksheets for executive review. | Automate nightly reporting that refreshes several pivot tables and archives the updated workbook as a single PDF file.
// AI Prompts: Write C# code using Aspose.Cells to add multiple pivot tables on different worksheets and export them to a single PDF while preserving each table’s layout. | Explain how PdfSaveOptions.ExportDocumentStructure and OnePagePerSheet affect PDF output for workbooks with pivot tables. | Troubleshoot why pivot tables lose formatting or appear blank after exporting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook with two data sheets, adds a pivot table on each sheet, refreshes them, configures PdfSaveOptions (ExportDocumentStructure = true, OnePagePerSheet = false) and saves all sheets as a single PDF while keeping each pivot's layout intact.
class ExportPivotTablesToPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ---------- First data sheet ----------
        Worksheet dataSheet1 = workbook.Worksheets[0];
        dataSheet1.Name = "Data1";
        dataSheet1.Cells["A1"].PutValue("Category");
        dataSheet1.Cells["B1"].PutValue("Amount");
        dataSheet1.Cells["A2"].PutValue("Food");
        dataSheet1.Cells["B2"].PutValue(120);
        dataSheet1.Cells["A3"].PutValue("Drink");
        dataSheet1.Cells["B3"].PutValue(80);
        dataSheet1.Cells["A4"].PutValue("Other");
        dataSheet1.Cells["B4"].PutValue(50);

        // Add first pivot table on a separate sheet
        Worksheet pivotSheet1 = workbook.Worksheets.Add("Pivot1");
        int ptIndex1 = pivotSheet1.PivotTables.Add("=Data1!A1:B4", "A1", "PivotTable1");
        PivotTable pivotTable1 = pivotSheet1.PivotTables[ptIndex1];
        pivotTable1.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable1.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivotTable1.ShowInTabularForm(); // preserve layout

        // ---------- Second data sheet ----------
        Worksheet dataSheet2 = workbook.Worksheets.Add("Data2");
        dataSheet2.Cells["A1"].PutValue("Product");
        dataSheet2.Cells["B1"].PutValue("Region");
        dataSheet2.Cells["C1"].PutValue("Sales");
        dataSheet2.Cells["A2"].PutValue("A");
        dataSheet2.Cells["B2"].PutValue("North");
        dataSheet2.Cells["C2"].PutValue(200);
        dataSheet2.Cells["A3"].PutValue("B");
        dataSheet2.Cells["B3"].PutValue("South");
        dataSheet2.Cells["C3"].PutValue(150);
        dataSheet2.Cells["A4"].PutValue("A");
        dataSheet2.Cells["B4"].PutValue("South");
        dataSheet2.Cells["C4"].PutValue(180);

        // Add second pivot table on another sheet
        Worksheet pivotSheet2 = workbook.Worksheets.Add("Pivot2");
        int ptIndex2 = pivotSheet2.PivotTables.Add("=Data2!A1:C4", "A1", "PivotTable2");
        PivotTable pivotTable2 = pivotSheet2.PivotTables[ptIndex2];
        pivotTable2.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable2.AddFieldToArea(PivotFieldType.Column, "Region");
        pivotTable2.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivotTable2.ShowInOutlineForm(); // preserve layout

        // Refresh all pivot tables to ensure they contain up‑to‑date data
        workbook.Worksheets.RefreshPivotTables();

        // Configure PDF save options to keep document structure (preserves each pivot layout)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true; // retain layout information
        pdfOptions.OnePagePerSheet = false;        // each sheet may span multiple pages as needed

        // Save the entire workbook (all sheets) to a single PDF file
        workbook.Save("PivotTables.pdf", pdfOptions);
    }
}
