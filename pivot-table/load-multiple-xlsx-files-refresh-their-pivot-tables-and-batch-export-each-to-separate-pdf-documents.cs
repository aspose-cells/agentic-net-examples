// Title: Refresh all pivot tables in multiple XLSX workbooks and export each workbook to a separate PDF using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a directory for .xlsx files, loads each workbook with Aspose.Cells, refreshes and recalculates every pivot table, then saves the workbook as a PDF with the same base name in an output folder. | Generate C# code that iterates through all worksheets and their pivot tables, calls PivotTable.RefreshData() and PivotTable.CalculateData(), and uses Workbook.Save to export each workbook to PDF, including folder creation and per‑file error handling.
// Common Searches: how to refresh pivot tables in all Excel files in a folder using Aspose.Cells C# | batch convert multiple XLSX files to PDF after updating pivot tables with .NET | C# loop through directory, refresh Excel pivot tables and save each workbook as PDF | Aspose.Cells bulk export of Excel workbooks to PDF after pivot refresh
// Tags: bulk Excel to PDF conversion Aspose.Cells | traverse workbook worksheets and pivot tables C# | automated folder processing for XLSX files .NET | error handling during batch workbook export | create output directory programmatically C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

namespace PivotTablePdfExport
{
    // The sample program enumerates every .xlsx file in a given input folder, loads each workbook with Aspose.Cells, refreshes and recalculates all pivot tables on every worksheet, and then saves the updated workbook as an individual PDF in a specified output folder, handling missing files, folder creation, and per‑file exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source XLSX files
            string sourceFolder = @"C:\InputXlsx";
            // Folder where the PDF files will be saved
            string outputFolder = @"C:\OutputPdf";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all XLSX files in the source folder
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                // Guard against missing file (should not happen with GetFiles, but added for safety)
                if (!File.Exists(xlsxPath))
                {
                    Console.WriteLine($"File not found: {xlsxPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(xlsxPath);

                    // Refresh all pivot tables in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        foreach (PivotTable pivotTable in sheet.PivotTables)
                        {
                            // Refresh the data source of the pivot table
                            pivotTable.RefreshData();
                            // Recalculate the pivot table after refresh
                            pivotTable.CalculateData();
                        }
                    }

                    // Build the output PDF file path (same name, .pdf extension)
                    string pdfFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                    Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {pdfFileName}");
                }
                catch (Exception ex)
                {
                    // Log any errors for the current file and continue processing others
                    Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("All files have been processed and exported to PDF.");
        }
    }
}
