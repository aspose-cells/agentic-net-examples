// Title: Merge Excel workbooks with Aspose.Cells for .NET and save as XLSX or PDF/A‑1b
// Description: Creates two workbooks, combines them using Workbook.Combine, saves the result as an XLSX file, and optionally exports a PDF/A‑1b version with embedded fonts and formula calculation via PdfSaveOptions.
// Keywords: Aspose.Cells | C# | .NET | Workbook.Combine | merge Excel files | save as XLSX | export to PDF | PdfSaveOptions | PDF/A-1b | embed fonts | calculate formulas
// Common Searches: Aspose.Cells merge two workbooks C# | Workbook.Combine example .NET | save merged workbook as XLSX with Aspose.Cells | export Excel workbook to PDF/A-1b using Aspose.Cells | PdfSaveOptions embed fonts C#
// Developer Intent: Combine multiple Excel workbooks into one file and produce XLSX and optional PDF/A‑1b outputs.
// Use Cases: Consolidate monthly reports from separate files into a single spreadsheet for distribution. | Create an archival PDF/A‑1b copy of a merged workbook to meet regulatory requirements. | Automate data aggregation pipelines that merge source workbooks and deliver both editable and read‑only formats.
// AI Prompts: Generate C# code that merges a list of Excel files with Aspose.Cells, renames each worksheet, and saves the result as XLSX and PDF/A‑2b. | Show how to configure PdfSaveOptions to embed custom fonts, set PDF/A‑2b compliance, and disable image compression when exporting a workbook. | Provide an example that loops through a directory, combines all workbooks using Workbook.Combine, and writes the final file in both XLSX and PDF formats.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

namespace AsposeCellsMergeAndSaveDemo
{
    // Creates two workbooks, combines them using Workbook.Combine, saves the result as an XLSX file, and optionally exports a PDF/A‑1b version with embedded fonts and formula calculation via PdfSaveOptions.
    public class Program
    {
        public static void Main()
        {
            // Create the first workbook and add sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "Source";
            sourceSheet.Cells["A1"].PutValue("Source Data");
            sourceSheet.Cells["B2"].PutValue(DateTime.Now);

            // Create the second workbook and add sample data
            Workbook destinationWorkbook = new Workbook();
            Worksheet destSheet = destinationWorkbook.Worksheets[0];
            destSheet.Name = "Destination";
            destSheet.Cells["A1"].PutValue("Destination Data");
            destSheet.Cells["C3"].PutValue(12345);

            // Combine the source workbook into the destination workbook
            // This uses the Workbook.Combine method (rule provided)
            destinationWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook as XLSX
            // Uses Workbook.Save(string, SaveFormat) rule
            string xlsxPath = "CombinedWorkbook.xlsx";
            destinationWorkbook.Save(xlsxPath, SaveFormat.Xlsx);
            Console.WriteLine($"Combined workbook saved as XLSX to '{xlsxPath}'.");

            // OPTIONAL: Save the same workbook as PDF for reporting
            // Uses Workbook.Save(string, PdfSaveOptions) rule
            string pdfPath = "CombinedWorkbook.pdf";
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example option: embed standard Windows fonts
                EmbedStandardWindowsFonts = true,
                // Example option: set compliance to PDF/A-1b
                Compliance = PdfCompliance.PdfA1b,
                // Example option: calculate formulas before saving
                CalculateFormula = true
            };
            destinationWorkbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"Combined workbook saved as PDF to '{pdfPath}'.");
        }
    }
}
