// Title: Aspose.Cells C# – Compare PDF number formatting under InvariantCulture vs French (fr‑FR)
// Description: This example creates a workbook with a numeric value formatted by "#,##0.00", switches Workbook.Settings.CultureInfo between InvariantCulture and French (fr‑FR), captures the formatted string for each locale, saves two PDFs (Invariant.pdf and French.pdf), and prints whether the displayed formats differ.
// Keywords: Aspose.Cells | C# | PDF export | CultureInfo | InvariantCulture | fr-FR | number formatting | localization | globalization | custom number format | Excel to PDF | regional settings
// Common Searches: Aspose.Cells PDF French number format | compare invariant and French culture in Aspose.Cells | C# set workbook CultureInfo for PDF export | Aspose.Cells localization PDF output | how to change number separators in exported PDF
// Developer Intent: Determine how switching the workbook's CultureInfo between invariant and French affects number formatting in the generated PDFs.
// Use Cases: Verify that financial PDFs show correct thousand and decimal separators for each target market. | Automate regression tests for PDF exports across multiple locales to catch formatting regressions. | Produce region‑specific invoices or reports where numeric values follow local conventions.
// AI Prompts: Generate C# code that loops through a list of cultures, saves a PDF for each with Aspose.Cells, and logs any differences in formatted numbers. | Explain the impact of Workbook.Settings.CultureInfo on number formatting during PDF conversion in Aspose.Cells. | Suggest an automated way to extract and compare formatted numbers inside the resulting PDFs for each culture.

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureComparison
{
    // This example creates a workbook with a numeric value formatted by "#,##0.00", switches Workbook.Settings.CultureInfo between InvariantCulture and French (fr‑FR), captures the formatted string for each locale, saves two PDFs (Invariant.pdf and French.pdf), and prints whether the displayed formats differ.
    class Program
    {
        static void Main()
        {
            // Create a workbook with a numeric value and a custom number format
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234567.89);

            // Apply a custom number format that uses group and decimal separators
            Style style = wb.CreateStyle();
            style.Custom = "#,##0.00";
            cell.SetStyle(style);

            // ---------- Invariant Culture ----------
            // Set workbook culture to invariant and capture the formatted string
            wb.Settings.CultureInfo = CultureInfo.InvariantCulture;
            string invariantFormatted = cell.StringValue;

            // Save PDF generated with invariant culture
            wb.Save("Invariant.pdf", SaveFormat.Pdf);

            // ---------- French Culture ----------
            // Change workbook culture to French (France) and capture the formatted string
            wb.Settings.CultureInfo = new CultureInfo("fr-FR");
            string frenchFormatted = cell.StringValue;

            // Save PDF generated with French culture
            wb.Save("French.pdf", SaveFormat.Pdf);

            // ---------- Comparison ----------
            Console.WriteLine($"Invariant format: {invariantFormatted}");
            Console.WriteLine($"French format:    {frenchFormatted}");
            Console.WriteLine($"Formats differ:   {invariantFormatted != frenchFormatted}");
        }
    }
}
