// Title: Load Large Excel Workbook via LightCells and Save as PDF – Aspose.Cells .NET
// Description: Shows how to stream a massive .xlsx file with Aspose.Cells LightCellsDataHandler, load it using LoadOptions, and directly save the workbook as a PDF without loading the entire file into memory.
// Keywords: Aspose.Cells | LightCells | C# | .NET | large Excel workbook | streaming load | memory‑efficient conversion | Excel to PDF | LightCellsDataHandler | save as PDF | large workbook PDF conversion
// Common Searches: Aspose.Cells LightCells load large workbook C# | Convert large Excel to PDF using LightCells .NET | Streaming Excel to PDF Aspose.Cells example | LightCellsDataHandler PDF conversion code | How to save big .xlsx as PDF without high memory usage
// Developer Intent: Load a massive Excel file with LightCells streaming mode and convert it to PDF in a single operation.
// Use Cases: Generate PDF reports from multi‑hundred‑megabyte Excel files on a web server. | Process client‑uploaded spreadsheets in a cloud function where memory is limited. | Integrate lightweight Excel‑to‑PDF conversion into batch jobs or micro‑services. | Apply a custom LightCellsDataHandler to skip hidden sheets before PDF export.
// AI Prompts: Write a LightCellsDataHandler that excludes hidden worksheets during PDF conversion. | Show how to configure PDF save options (page size, orientation, compression) after loading with LightCells. | Provide robust error handling for missing or corrupted source files in LightCells streaming conversion. | Explain how to monitor memory usage while converting a large workbook with LightCells. | Generate code to convert multiple large workbooks to PDFs in parallel using LightCells.

using System;
using Aspose.Cells;

namespace LightCellsPdfConversion
{
    // Custom handler that processes cells in a lightweight streaming manner.
    // For this example we simply allow all sheets, rows, and cells to be processed
    // without performing any additional logic.
    // Shows how to stream a massive .xlsx file with Aspose.Cells LightCellsDataHandler, load it using LoadOptions, and directly save the workbook as a PDF without loading the entire file into memory.
    public class SimpleLightCellsDataHandler : LightCellsDataHandler
    {
        // Called when a worksheet is about to be processed.
        public bool StartSheet(Worksheet sheet)
        {
            // Return true to continue processing this sheet.
            return true;
        }

        // Called before processing a specific row.
        public bool StartRow(int rowIndex)
        {
            // Return true to process the row.
            return true;
        }

        // Called after a row has been read; can be used to inspect row data.
        public bool ProcessRow(Row row)
        {
            // No custom processing needed; continue.
            return true;
        }

        // Called before processing a specific cell in the current row.
        public bool StartCell(int columnIndex)
        {
            // Return true to process the cell.
            return true;
        }

        // Called after a cell has been read; can be used to inspect cell data.
        public bool ProcessCell(Cell cell)
        {
            // No custom processing needed; continue.
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the large Excel file to be loaded.
            string sourcePath = "LargeWorkbook.xlsx";

            // Path where the resulting PDF will be saved.
            string pdfPath = "LargeWorkbook.pdf";

            // Create load options and assign the LightCellsDataHandler.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new SimpleLightCellsDataHandler();

            // Load the workbook using the LightCells streaming mode.
            // This uses the constructor: Workbook(string file, LoadOptions loadOptions)
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the loaded workbook directly to PDF.
            // This uses the method: Save(string fileName, SaveFormat saveFormat)
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine("Workbook loaded with LightCells API and saved to PDF successfully.");
        }
    }
}
