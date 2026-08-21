// Title: Save a merged Aspose.Cells workbook as XLSX with optional PDF export (C#)
// Description: Demonstrates how to persist a merged Aspose.Cells Workbook in the native XLSX format and, when required, generate a PDF copy for reporting. The example includes basic cell writing, exception handling, and a reusable method that accepts a Workbook and an export flag.
// Keywords: Aspose.Cells C# save workbook XLSX | Aspose.Cells export to PDF | save merged workbook Aspose | Aspose.Cells SaveFormat Xlsx | Aspose.Cells SaveFormat Pdf | C# Aspose.Cells multiple output formats | generate Excel and PDF with Aspose.Cells
// Common Searches: C# Aspose.Cells save workbook as xlsx and pdf | How to export a merged workbook to PDF using Aspose.Cells | Aspose.Cells SaveFormat Xlsx example | Save Aspose.Cells workbook to multiple formats | Aspose.Cells generate PDF report from workbook
// Developer Intent: Persist a merged workbook in XLSX format and optionally create a PDF version for distribution or reporting.
// Use Cases: Consolidate several source workbooks, save the final Excel file for downstream processing, and provide a PDF snapshot for stakeholders. | Automate nightly data merges where the .xlsx file is archived and a PDF is emailed to a distribution list. | Offer a downloadable Excel report in a web app while simultaneously showing a PDF preview generated from the same workbook.
// AI Prompts: Write a C# method that takes an Aspose.Cells Workbook and saves it as XLSX and PDF with customizable file names and output folder. | Show how to implement robust error handling when saving an Aspose.Cells workbook to multiple formats and log any exceptions. | Provide sample code that merges multiple workbooks using Aspose.Cells and then calls a helper to export both XLSX and PDF versions.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to persist a merged Aspose.Cells Workbook in the native XLSX format and, when required, generate a PDF copy for reporting. The example includes basic cell writing, exception handling, and a reusable method that accepts a Workbook and an export flag.
    public class SaveMergedWorkbook
    {
        // Saves the provided merged workbook as XLSX and optionally as PDF.
        public static void Run(Workbook mergedWorkbook, bool exportPdf = true)
        {
            try
            {
                // Save the workbook in the native Excel format (XLSX).
                mergedWorkbook.Save("MergedOutput.xlsx", SaveFormat.Xlsx);

                // If a PDF copy is required, save the workbook as PDF.
                if (exportPdf)
                {
                    mergedWorkbook.Save("MergedOutput.pdf", SaveFormat.Pdf);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a sample workbook (could be merged from other sources)
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");

                // Call the save method
                SaveMergedWorkbook.Run(wb, exportPdf: true);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
