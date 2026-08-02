// Title: Export Pivot Table to PDF with Layout & Accessibility using Aspose.Cells for .NET
// Description: C# example that creates a workbook, builds a pivot table, refreshes its data, and saves the entire workbook as a PDF. The PdfSaveOptions are configured to keep the original layout, column widths, and document structure tags, and to fit each worksheet on a single page.
// Keywords: Aspose.Cells export pivot table PDF | C# PdfSaveOptions layout | preserve pivot formatting PDF | one page per sheet Aspose | PDF accessibility tags Aspose.Cells | export workbook to PDF .NET | pivot table PDF export code
// Common Searches: How to export a pivot table to PDF with Aspose.Cells C# | Aspose.Cells keep column widths when saving PDF | PDF export options for pivot tables Aspose .NET | Save workbook as PDF with accessibility tags Aspose.Cells | One page per sheet PDF Aspose.Cells example
// Developer Intent: Generate a PDF from a workbook that contains a pivot table, ensuring the table’s layout, column widths, and accessibility tags are retained.
// Use Cases: Create printable sales reports that include a formatted pivot table. | Deliver PDF dashboards to stakeholders where each sheet fits on a single page. | Produce accessible PDFs (tagged for screen readers) from Excel workbooks with pivot tables.
// AI Prompts: Show C# code using Aspose.Cells to export a workbook with a pivot table to PDF while preserving layout and column widths. | Explain how to configure PdfSaveOptions for accessibility tags and one‑page‑per‑sheet output in Aspose.Cells. | Detail the steps to refresh, calculate, and save a pivot‑driven workbook as a PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that creates a workbook, builds a pivot table, refreshes its data, and saves the entire workbook as a PDF. The PdfSaveOptions are configured to keep the original layout, column widths, and document structure tags, and to fit each worksheet on a single page.
class ExportPivotToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet for source data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample data for the pivot table
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Product");
        dataSheet.Cells["C1"].PutValue("Sales");

        dataSheet.Cells["A2"].PutValue("2023");
        dataSheet.Cells["B2"].PutValue("Apple");
        dataSheet.Cells["C2"].PutValue(1200);

        dataSheet.Cells["A3"].PutValue("2023");
        dataSheet.Cells["B3"].PutValue("Banana");
        dataSheet.Cells["C3"].PutValue(1500);

        dataSheet.Cells["A4"].PutValue("2023");
        dataSheet.Cells["B4"].PutValue("Apple");
        dataSheet.Cells["C4"].PutValue(800);

        // Add a new worksheet to host the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table (source range = Data!A1:C4, destination = A3)
        int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C4", "A3", "SalesPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure pivot fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Layout the pivot table in tabular form for better appearance
        pivotTable.ShowInTabularForm();

        // Refresh and calculate the pivot data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Refresh any other pivot tables in the workbook (good practice)
        workbook.Worksheets.RefreshPivotTables();

        // Set PDF save options to preserve layout and formatting
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export the document structure (tags) to keep accessibility information
            ExportDocumentStructure = true,

            // Fit each worksheet onto a single PDF page while keeping column widths
            OnePagePerSheet = true,
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook (including the pivot table) as a PDF file
        workbook.Save("PivotTableOutput.pdf", pdfOptions);
    }
}
