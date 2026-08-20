// Title: Export Pivot Table to PDF with Layout Preservation using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample sales data, builds a pivot table, sets tabular layout, refreshes the cache, and saves the workbook as a PDF with PdfSaveOptions that keep column widths, one‑page‑per‑sheet layout, and document structure tags.
// Keywords: Aspose.Cells PDF export | pivot table to PDF | C# Aspose.Cells PdfSaveOptions | preserve pivot formatting PDF | OnePagePerSheet Aspose | export document structure Aspose | Excel pivot PDF .NET
// Common Searches: export Excel pivot table to PDF Aspose.Cells | keep pivot table layout when saving as PDF C# | PdfSaveOptions preserve column widths Aspose | how to export workbook with pivot to PDF .NET | Aspose.Cells generate PDF report with pivot
// Developer Intent: Generate a PDF from a workbook that contains a pivot table while retaining the table’s layout, column widths, and accessibility tags.
// Use Cases: Produce a printable sales summary PDF that includes a pivot table grouped by category and product. | Automate PDF creation for financial dashboards where pivot tables must maintain their exact visual structure. | Create accessible PDF reports with document‑structure tags for compliance, preserving the original Excel pivot formatting.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook containing a pivot table to PDF, preserving tabular layout and using PdfSaveOptions for one page per sheet. | Explain how to configure PdfSaveOptions to keep column widths and export document structure when saving a pivot table to PDF with Aspose.Cells. | Show the steps to refresh and calculate a pivot table before converting the workbook to PDF in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotPdfExport
{
    // Creates a workbook, adds sample sales data, builds a pivot table, sets tabular layout, refreshes the cache, and saves the workbook as a PDF with PdfSaveOptions that keep column widths, one‑page‑per‑sheet layout, and document structure tags.
    public class ExportPivotTableToPdf
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (data sheet)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(800);

            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(600);

            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue("Tomato");
            dataSheet.Cells["C5"].PutValue(900);

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table on the pivot sheet
            // Source range: Data!A1:C5, Destination cell: A3, Pivot table name: SalesPivot
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");      // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");    // Column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");       // Data field

            // Layout the pivot table in tabular form to preserve formatting
            pivotTable.ShowInTabularForm();

            // Refresh the pivot cache and calculate data (correct API usage)
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Set PDF save options to keep layout and document structure
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export the document structure (tags) for better accessibility
                ExportDocumentStructure = true,
                // Fit each worksheet onto a single page while preserving column widths
                OnePagePerSheet = true,
                AllColumnsInOnePagePerSheet = true
            };

            // Save the workbook (including the pivot table) as a PDF file
            try
            {
                workbook.Save("PivotTableExport.pdf", pdfOptions);
                Console.WriteLine("Workbook with pivot table exported to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save PDF: {ex.Message}");
            }
        }
    }
}
