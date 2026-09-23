// Title: Compare PDF output of an Excel workbook when exported with InvariantCulture versus French (fr-FR) culture using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets Workbook.Settings.CultureInfo to CultureInfo.InvariantCulture, saves the workbook as PDF, then changes CultureInfo to fr-FR, saves a second PDF, and compares the two PDF files byte‑by‑byte to determine if any formatting differences exist. | Extend the sample to output a log that records each numeric cell's formatted value in the invariant‑culture PDF and the French‑culture PDF, highlighting differences such as decimal separators or grouping symbols.
// Common Searches: aspnet compare pdf files generated from Excel with different CultureInfo settings | Aspose.Cells export to PDF using invariant culture vs French locale | detect number format changes in PDF when changing workbook culture in C# | how to set workbook.Settings.CultureInfo for PDF conversion in Aspose.Cells | byte array comparison of two PDFs created by Aspose.Cells
// Tags: Aspose.Cells culture‑specific PDF generation | C# workbook.Settings.CultureInfo configuration | byte‑wise PDF difference detection Aspose.Cells | locale number formatting in Excel to PDF | French (fr-FR) PDF output with Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, saves it as PDF twice—first using InvariantCulture and then using French (fr-FR) culture—reads both PDFs into byte arrays, and performs a byte‑wise comparison to report whether the outputs differ, illustrating locale‑dependent number formatting.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(excelPath);

        // ---------- Invariant Culture ----------
        // Set culture to invariant for number formatting
        workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;

        // Save as PDF using invariant culture (lifecycle rule: save)
        string pdfInvariantPath = "output_invariant.pdf";
        workbook.Save(pdfInvariantPath, SaveFormat.Pdf);

        // ---------- French Culture ----------
        // Change culture to French (France) for number formatting
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Save as PDF using French culture
        string pdfFrenchPath = "output_french.pdf";
        workbook.Save(pdfFrenchPath, SaveFormat.Pdf);

        // ---------- Comparison ----------
        // Load both PDF files into byte arrays
        byte[] invariantBytes = File.ReadAllBytes(pdfInvariantPath);
        byte[] frenchBytes = File.ReadAllBytes(pdfFrenchPath);

        // Simple byte‑wise comparison to detect differences
        bool areIdentical = invariantBytes.Length == frenchBytes.Length;
        if (areIdentical)
        {
            for (int i = 0; i < invariantBytes.Length; i++)
            {
                if (invariantBytes[i] != frenchBytes[i])
                {
                    areIdentical = false;
                    break;
                }
            }
        }

        // Output the result
        Console.WriteLine("PDFs are {0}.", areIdentical ? "identical" : "different");
    }
}
