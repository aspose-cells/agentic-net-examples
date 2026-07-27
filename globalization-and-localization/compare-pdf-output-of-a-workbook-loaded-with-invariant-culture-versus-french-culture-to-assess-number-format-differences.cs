using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCultureComparison
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a sample workbook with numeric data
            Workbook sourceWb = new Workbook();
            Worksheet sheet = sourceWb.Worksheets[0];
            // Put a number that will be formatted differently in French culture
            sheet.Cells["A1"].PutValue(1234567.89);
            // Apply a number format with two decimal places
            Style style = sourceWb.CreateStyle();
            style.Custom = "#,##0.00";
            sheet.Cells["A1"].SetStyle(style);
            // Save the source workbook (XLSX) – this is the template for both loads
            sourceWb.Save("sample.xlsx", SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // Step 2: Load with InvariantCulture and export to PDF
            LoadOptions invariantOptions = new LoadOptions(LoadFormat.Xlsx);
            invariantOptions.CultureInfo = CultureInfo.InvariantCulture; // invariant culture
            Workbook invariantWb = new Workbook("sample.xlsx", invariantOptions);
            // Export to PDF
            invariantWb.Save("output_invariant.pdf", SaveFormat.Pdf);
            // Capture the formatted string of the cell for later comparison
            string invariantFormatted = invariantWb.Worksheets[0].Cells["A1"].StringValue;

            // -----------------------------------------------------------------
            // Step 3: Load with French culture and export to PDF
            LoadOptions frenchOptions = new LoadOptions(LoadFormat.Xlsx);
            frenchOptions.CultureInfo = new CultureInfo("fr-FR"); // French (France) culture
            Workbook frenchWb = new Workbook("sample.xlsx", frenchOptions);
            // Export to PDF
            frenchWb.Save("output_french.pdf", SaveFormat.Pdf);
            // Capture the formatted string of the cell for later comparison
            string frenchFormatted = frenchWb.Worksheets[0].Cells["A1"].StringValue;

            // -----------------------------------------------------------------
            // Step 4: Compare the formatted values and output the differences
            Console.WriteLine("Formatted value with InvariantCulture: " + invariantFormatted);
            Console.WriteLine("Formatted value with French culture:    " + frenchFormatted);

            if (invariantFormatted == frenchFormatted)
            {
                Console.WriteLine("No difference in number formatting between the two cultures.");
            }
            else
            {
                Console.WriteLine("Difference detected:");
                Console.WriteLine($" - Invariant uses: '{invariantFormatted}'");
                Console.WriteLine($" - French uses:    '{frenchFormatted}'");
            }

            // Note: The generated PDFs (output_invariant.pdf and output_french.pdf)
            // can be manually inspected to see the visual differences in number formatting.
        }
    }
}