// Title: Export a smart‑marker driven Excel workbook with charts to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel template containing smart markers and embedded charts, binds a DataTable as the data source, processes the smart markers, and saves the workbook as a PDF while keeping the charts intact. | Show how to configure PdfSaveOptions with ExportDocumentStructure = true and CalculateFormula = true to retain chart graphics and evaluate formulas when exporting a smart‑marker workbook to PDF using Aspose.Cells. | Explain the steps to use WorkbookDesigner to fill a smart‑marker template, update chart data ranges, and then generate a PDF file that includes all graphics.
// Common Searches: how to keep charts when exporting an Aspose.Cells workbook with smart markers to PDF | C# Aspose.Cells PdfSaveOptions ExportDocumentStructure example for smart markers | populate Excel smart marker template from DataTable and save as PDF with graphics | export smart marker driven workbook to PDF preserving document structure Aspose.Cells | convert Excel file with charts to PDF using Aspose.Cells without losing images
// Tags: smart marker workbook PDF export preserving charts | Aspose.Cells PdfSaveOptions ExportDocumentStructure usage | populate smart marker template from DataTable C# | convert Excel with charts to PDF Aspose.Cells | process smart markers before PDF generation

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel template that contains smart markers and charts, fills it with a DataTable via WorkbookDesigner, processes the markers to update chart data, configures PdfSaveOptions to retain document structure and calculate formulas, and saves the result as a PDF with all graphics preserved.
class ExportWorkbookToPdf
{
    static void Main()
    {
        // Load a workbook that contains smart markers and charts.
        // The template can be created beforehand with markers like &amp;=Products.Name etc.
        Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

        // -------------------- Prepare data source --------------------
        // Example data table that matches the smart marker names used in the template.
        DataTable products = new DataTable("Products");
        products.Columns.Add("Name", typeof(string));
        products.Columns.Add("Quantity", typeof(int));

        // Populate the table with sample data.
        products.Rows.Add("Apple", 50);
        products.Rows.Add("Banana", 30);
        products.Rows.Add("Cherry", 20);

        // -------------------- Process smart markers --------------------
        // Attach the data source to the workbook designer and populate the template.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(products);
        designer.Process(); // fills cells, updates chart data ranges, etc.

        // -------------------- Configure PDF save options --------------------
        // ExportDocumentStructure ensures that charts and other graphics are retained.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,
            CalculateFormula = true // calculate any formulas before saving
        };

        // -------------------- Save as PDF --------------------
        // Use the provided Save method that accepts SaveOptions.
        workbook.Save("Result.pdf", pdfOptions);
    }
}
