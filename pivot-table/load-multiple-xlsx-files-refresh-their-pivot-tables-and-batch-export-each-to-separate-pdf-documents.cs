// Title: Refresh pivot tables in multiple XLSX workbooks and export each workbook to a PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a list of .xlsx files, calls Worksheets.RefreshPivotTables on each workbook, and saves the result as a PDF using ConversionUtility.Convert. | Generate a batch routine in .NET that iterates over Excel files, refreshes all pivot tables, and creates separate PDF files for each workbook with Aspose.Cells. | Create a console application that validates file existence, refreshes pivot caches, and converts refreshed workbooks to PDFs in a single pass.
// Common Searches: aspnet refresh all pivot tables in a folder of Excel files using Aspose.Cells | batch convert refreshed XLSX workbooks to PDF with C# Aspose.Cells | how to loop through multiple Excel workbooks and update pivot tables programmatically | Aspose.Cells Worksheets.RefreshPivotTables example for bulk processing | convert Excel workbook to PDF after pivot refresh using ConversionUtility in .NET
// Tags: bulk refresh pivot tables Aspose.Cells | pivot table refresh and PDF export .NET | Worksheets.RefreshPivotTables batch processing | ConversionUtility.Convert PDF generation | process multiple XLSX files Aspose.Cells | export refreshed workbook to PDF C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace PivotRefreshAndPdfExport
{
    // The example iterates over a collection of XLSX files, loads each workbook with Aspose.Cells, refreshes all pivot tables via Worksheets.RefreshPivotTables, saves the workbook, and then converts the refreshed file to a separate PDF using ConversionUtility.Convert.
    class Program
    {
        static void Main()
        {
            // List of source Excel files (XLSX) to process
            List<string> sourceFiles = new List<string>
            {
                "Report1.xlsx",
                "Report2.xlsx",
                "Report3.xlsx"
                // Add more file paths as needed
            };

            // Process each file
            foreach (string sourcePath in sourceFiles)
            {
                // Ensure the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Refresh all pivot tables in the workbook
                workbook.Worksheets.RefreshPivotTables();

                // Save the refreshed workbook back to the same file (or to a temp file)
                // Using the standard Save method as there is no specific rule for saving.
                workbook.Save(sourcePath);

                // Determine the PDF output path (same name with .pdf extension)
                string pdfPath = Path.ChangeExtension(sourcePath, ".pdf");

                // Convert the refreshed Excel file to PDF using the provided ConversionUtility.Convert rule
                ConversionUtility.Convert(sourcePath, pdfPath);

                Console.WriteLine($"Converted '{sourcePath}' to PDF '{pdfPath}'.");
            }

            Console.WriteLine("All files processed.");
        }
    }
}
