// Title: Export a workbook with a pivot table to PDF while preserving layout using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds a pivot table, switches it to tabular form, and saves the workbook as a PDF using Aspose.Cells with layout‑preserving settings. | Demonstrate how to set PDF export options (single‑page per sheet, fit all columns, retain document structure) to render a pivot‑table workbook correctly in Aspose.Cells.
// Common Searches: asp.net export pivot table to pdf with aspose.cells preserving formatting | c# keep pivot table layout when saving workbook as pdf using aspose | pdfsaveoptions onepagepersheet allcolumnsononepage for pivot tables aspose.cells | export workbook containing pivot table to single page pdf c# aspose.cells example | aspose.cells pdf export pivot table tabular form layout
// Tags: Aspose.Cells export workbook to PDF | pivot table PDF layout preservation | PdfSaveOptions OnePagePerSheet setting | tabular form pivot table rendering | C# Aspose.Cells pivot table export

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, adds sample data, defines a pivot table on a separate sheet, sets the pivot to tabular form, refreshes and calculates it, configures PdfSaveOptions to keep the layout (one page per sheet, all columns on one page, export document structure), and saves the result as PivotTableOutput.pdf.
public class ExportPivotTableToPdf
{
    public static void Main()
    {
        try
        {
            Run();
            Console.WriteLine("Pivot table exported to PDF successfully.");
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

        // Add sample data to the first worksheet
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        dataSheet.Cells["A1"].PutValue("Year");
        dataSheet.Cells["B1"].PutValue("Product");
        dataSheet.Cells["C1"].PutValue("Sales");

        dataSheet.Cells["A2"].PutValue(2023);
        dataSheet.Cells["B2"].PutValue("Apple");
        dataSheet.Cells["C2"].PutValue(1200);

        dataSheet.Cells["A3"].PutValue(2023);
        dataSheet.Cells["B3"].PutValue("Banana");
        dataSheet.Cells["C3"].PutValue(1500);

        dataSheet.Cells["A4"].PutValue(2023);
        dataSheet.Cells["B4"].PutValue("Apple");
        dataSheet.Cells["C4"].PutValue(800);

        // Create a worksheet to host the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

        // Add a pivot table based on the data range
        int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C4", "A3", "SalesPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure pivot fields: Year as row, Product as column, Sales as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Year");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Layout the pivot table in tabular form for better PDF rendering
        pivotTable.ShowInTabularForm();

        // Refresh and calculate the pivot data using the correct API
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Set PDF save options to preserve layout and formatting
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true,                 // each sheet on a single page
            AllColumnsInOnePagePerSheet = true,    // fit all columns on that page
            ExportDocumentStructure = true         // retain document structure
        };

        // Export the workbook (including the pivot table) to PDF
        string outputPath = "PivotTableOutput.pdf";

        try
        {
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception saveEx)
        {
            Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
            throw;
        }
    }
}
