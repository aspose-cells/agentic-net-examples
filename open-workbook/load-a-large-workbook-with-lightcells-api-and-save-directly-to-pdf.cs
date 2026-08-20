// Title: Load a Large Excel Workbook with LightCells API and Save Directly to PDF (C#)
// Description: Demonstrates how to stream‑load a massive .xlsx file using Aspose.Cells LightCellsDataHandler, then convert and save it to PDF in one step, minimizing memory consumption.
// Keywords: Aspose.Cells LightCells | C# load large workbook | streaming Excel to PDF | memory‑efficient Excel conversion | LightCellsDataHandler PDF | large .xlsx to PDF | Aspose.Cells .NET PDF export | GitHub Aspose.Cells example
// Common Searches: load large excel file with LightCells C# | convert big .xlsx to pdf using Aspose.Cells | streaming workbook load Aspose.Cells LightCellsDataHandler | memory efficient excel to pdf conversion .NET | Aspose.Cells LightCells PDF export example
// Developer Intent: Stream a huge Excel workbook and export it to PDF without loading the entire file into memory.
// Use Cases: Generate PDF reports from multi‑gigabyte Excel files in a web service. | Batch‑process enterprise spreadsheets to PDF while keeping server RAM usage low. | Create on‑the‑fly PDF previews of user‑uploaded workbooks in cloud applications.
// AI Prompts: Write a LightCellsDataHandler that skips hidden rows and columns during PDF conversion. | Show how to set PDF save options (image quality, PDF/A compliance) when using LightCells streaming. | Explain error handling for corrupted large workbooks while still attempting PDF export.

using System;
using Aspose.Cells;

// Demonstrates how to stream‑load a massive .xlsx file using Aspose.Cells LightCellsDataHandler, then convert and save it to PDF in one step, minimizing memory consumption.
class LightCellsPdfConversion
{
    static void Main()
    {
        // Paths for source Excel file and destination PDF file
        string sourcePath = "LargeWorkbook.xlsx";
        string pdfPath = "LargeWorkbook.pdf";

        // Create a LightCellsDataHandler that simply processes all sheets, rows, and cells
        LightCellsDataHandler handler = new SimpleHandler();

        // Configure LoadOptions to use the LightCellsDataHandler for streaming load
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = handler;

        // Load the large workbook using LightCells API (streaming mode)
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Save the loaded workbook directly to PDF format
        workbook.Save(pdfPath, SaveFormat.Pdf);
    }

    // Minimal implementation of LightCellsDataHandler that accepts everything
    class SimpleHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet) => true;          // Process every worksheet
        public bool StartRow(int rowIndex) => true;              // Process every row
        public bool ProcessRow(Row row) => true;                 // No custom row processing
        public bool StartCell(int columnIndex) => true;          // Process every cell
        public bool ProcessCell(Cell cell) => true;              // No custom cell processing
    }
}
