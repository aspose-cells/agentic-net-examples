// Title: Load Workbook with LightCells API, Strip Charts, and Export to PDF – Aspose.Cells for .NET
// Description: Demonstrates how to open an XLSX file using Aspose.Cells LightCells (MemoryPreference), clear every worksheet's chart collection, and save the result as a PDF with default PdfSaveOptions. Includes file‑existence check and exception handling for robust server‑side processing.
// Keywords: Aspose.Cells LightCells load workbook | exclude charts Excel to PDF | MemoryPreference PDF conversion .NET | remove worksheet charts Aspose | light memory mode Excel PDF export
// Common Searches: load xlsx with LightCells and omit charts | Aspose.Cells .NET export workbook to PDF without charts | how to clear charts before PDF conversion using Aspose | LightCells MemoryPreference reduce memory when converting to PDF
// Developer Intent: Open an Excel file with LightCells, delete all chart objects, and generate a PDF.
// Use Cases: Create lightweight PDFs of large financial spreadsheets where visual charts are unnecessary. | Automate batch processing of user‑uploaded workbooks on a cloud server, stripping charts to meet compliance or size constraints. | Produce printable reports that contain only tabular data, improving rendering speed and reducing PDF file size.
// AI Prompts: Generate C# code that uses Aspose.Cells LightCells (MemoryPreference) to load an XLSX file, remove every chart from each worksheet, and save the workbook as a PDF. | Show an example of configuring LoadOptions for LightCells, clearing the Charts collection, and handling missing files or runtime errors during PDF export. | Explain best practices for converting large Excel files to PDF with Aspose.Cells while excluding charts to minimize memory usage.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsLightCellsExample
{
    // Demonstrates how to open an XLSX file using Aspose.Cells LightCells (MemoryPreference), clear every worksheet's chart collection, and save the result as a PDF with default PdfSaveOptions. Includes file‑existence check and exception handling for robust server‑side processing.
    public class ExcludeChartsAndSavePdf
    {
        public static void Run()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Configure LoadOptions to use LightCells memory mode
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    MemorySetting = MemorySetting.MemoryPreference
                };

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Optional: remove charts if they were loaded (safety net)
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.Charts.Clear();
                }

                // Create PDF save options (default configuration)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook to PDF; charts are omitted
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Workbook loaded without charts and saved to PDF at: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
