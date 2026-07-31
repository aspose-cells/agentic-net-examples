// Title: Batch Refresh Pivot Tables and Export XLSX to PDF using Aspose.Cells (.NET)
// Description: Scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, refreshes every pivot table, saves the updated file, and converts it to a PDF via ConversionUtility. Includes per‑file error handling and console logging.
// Keywords: Aspose.Cells | C# pivot table refresh | batch Excel to PDF | refresh all pivot tables | convert XLSX to PDF .NET | process multiple workbooks | ConversionUtility example | automate Excel reporting
// Common Searches: Aspose.Cells batch refresh pivot tables | Convert multiple Excel files to PDF C# | Refresh pivot tables in all worksheets programmatically | How to export refreshed XLSX files to PDF with Aspose | Automate Excel to PDF conversion for a folder
// Developer Intent: Refresh every pivot table in each Excel workbook within a directory and generate a matching PDF file for each workbook.
// Use Cases: Nightly job that updates financial dashboards and archives them as PDFs. | Bulk processing of client‑submitted reports: refresh data, then deliver PDF versions. | Server‑side service that receives an XLSX, refreshes its pivots, and returns a PDF response.
// AI Prompts: Generate C# code that iterates over all .xlsx files in a folder, refreshes all pivot tables with Aspose.Cells, and saves each workbook. | Show how to convert a refreshed workbook to PDF using Aspose.Cells ConversionUtility with proper exception handling. | Create a logging snippet that records the input file name, success status, and output PDF path during batch export.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace PivotTableBatchExport
{
    // Scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, refreshes every pivot table, saves the updated file, and converts it to a PDF via ConversionUtility. Includes per‑file error handling and console logging.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source XLSX files
            string sourceFolder = @"C:\InputExcels";

            // Ensure the folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Get all XLSX files in the folder
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Refresh all pivot tables in all worksheets
                    workbook.Worksheets.RefreshPivotTables();

                    // Save the refreshed workbook (overwrites the original file)
                    workbook.Save(excelPath);

                    // Determine the PDF output path (same name, .pdf extension)
                    string pdfPath = Path.ChangeExtension(excelPath, ".pdf");

                    // Convert the refreshed Excel file to PDF
                    ConversionUtility.Convert(excelPath, pdfPath);

                    Console.WriteLine($"Successfully exported '{Path.GetFileName(excelPath)}' to PDF.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
    }
}
