// Title: Refresh Pivot Tables in Multiple XLSX Files and Export Each to PDF with Aspose.Cells for .NET
// Description: A C# console app that loops through a list of Excel workbooks, validates each file, refreshes every pivot table using Worksheets.RefreshPivotTables(), saves the changes, and converts the updated workbook to a same‑named PDF via ConversionUtility.
// Keywords: Aspose.Cells | C# pivot table refresh | batch Excel to PDF | RefreshPivotTables | ConversionUtility | multiple XLSX to PDF | automated report generation
// Common Searches: aspocells refresh pivot tables batch | c# convert multiple excel files to pdf | update pivot caches before pdf export | automate excel pivot refresh and pdf conversion | aspocells batch processing example
// Developer Intent: Update all pivot tables in each supplied workbook and produce an individual PDF for every file.
// Use Cases: Nightly automation that refreshes pivot‑driven dashboards and distributes PDFs to business users. | Server‑side service that receives uploaded Excel reports, synchronizes pivot data, and returns PDF versions. | Bulk migration of a folder of legacy XLSX reports into PDF format while ensuring the latest calculations are reflected.
// AI Prompts: Generate C# code that iterates over a collection of Excel paths, calls Worksheets.RefreshPivotTables() for each workbook, and saves the result as PDF using ConversionUtility. | Provide a robust error‑handling pattern for missing files, read/write permissions, and logging conversion outcomes in a batch Excel‑to‑PDF routine with Aspose.Cells. | Compare the performance and feature differences between ConversionUtility.Convert and Workbook.SaveAsPdf when exporting many workbooks in a loop.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace PivotTableBatchPdfExport
{
    // A C# console app that loops through a list of Excel workbooks, validates each file, refreshes every pivot table using Worksheets.RefreshPivotTables(), saves the changes, and converts the updated workbook to a same‑named PDF via ConversionUtility.
    class Program
    {
        static void Main()
        {
            // List of Excel files to process
            List<string> excelFiles = new List<string>
            {
                "Report1.xlsx",
                "Report2.xlsx",
                "Report3.xlsx"
                // Add more file paths as needed
            };

            foreach (string excelPath in excelFiles)
            {
                // Verify the source file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Refresh all pivot tables in the workbook
                workbook.Worksheets.RefreshPivotTables();

                // Save the refreshed workbook back to the same file (or to a temp file)
                workbook.Save(excelPath);

                // Determine PDF output path (same name with .pdf extension)
                string pdfPath = Path.ChangeExtension(excelPath, ".pdf");

                // Convert the refreshed Excel file to PDF
                ConversionUtility.Convert(excelPath, pdfPath);

                Console.WriteLine($"Converted '{excelPath}' to PDF: '{pdfPath}'");
            }
        }
    }
}
