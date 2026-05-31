using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsCulturePdfComparison
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a sample workbook with numeric data and a custom number format.
            Workbook sampleWorkbook = new Workbook();
            Worksheet sheet = sampleWorkbook.Worksheets[0];
            // Populate some numbers.
            sheet.Cells["A1"].PutValue(1234.56);
            sheet.Cells["A2"].PutValue(7890.12);
            sheet.Cells["A3"].PutValue(3456.78);
            // Apply a number format that uses group and decimal separators.
            Style numberStyle = sampleWorkbook.CreateStyle();
            numberStyle.Custom = "#,##0.00";
            sheet.Cells["A1"].SetStyle(numberStyle);
            sheet.Cells["A2"].SetStyle(numberStyle);
            sheet.Cells["A3"].SetStyle(numberStyle);

            // Save the workbook to a temporary XLSX file (used as the source for both loads).
            string sourcePath = "sample.xlsx";
            sampleWorkbook.Save(sourcePath, SaveFormat.Xlsx);

            // Step 2: Load the workbook with InvariantCulture and export to PDF.
            LoadOptions invariantOptions = new LoadOptions(LoadFormat.Xlsx);
            invariantOptions.CultureInfo = CultureInfo.InvariantCulture;
            Workbook wbInvariant = new Workbook(sourcePath, invariantOptions);
            string pdfInvariantPath = "output_invariant.pdf";
            wbInvariant.Save(pdfInvariantPath, SaveFormat.Pdf);

            // Step 3: Load the workbook with French culture (fr-FR) and export to PDF.
            LoadOptions frenchOptions = new LoadOptions(LoadFormat.Xlsx);
            frenchOptions.CultureInfo = new CultureInfo("fr-FR");
            Workbook wbFrench = new Workbook(sourcePath, frenchOptions);
            string pdfFrenchPath = "output_french.pdf";
            wbFrench.Save(pdfFrenchPath, SaveFormat.Pdf);

            // Step 4: Compare the two PDF files byte‑by‑byte.
            byte[] pdfInvariantBytes = File.ReadAllBytes(pdfInvariantPath);
            byte[] pdfFrenchBytes = File.ReadAllBytes(pdfFrenchPath);
            bool pdfsAreIdentical = pdfInvariantBytes.SequenceEqual(pdfFrenchBytes);

            Console.WriteLine($"PDFs are {(pdfsAreIdentical ? "identical" : "different")}.");

            // Optional: Show a simple metric of difference (number of differing bytes).
            if (!pdfsAreIdentical)
            {
                int diffCount = pdfInvariantBytes.Zip(pdfFrenchBytes, (b1, b2) => b1 == b2 ? 0 : 1).Sum();
                // Account for length differences.
                diffCount += Math.Abs(pdfInvariantBytes.Length - pdfFrenchBytes.Length);
                Console.WriteLine($"Number of differing bytes: {diffCount}");
            }

            // Cleanup temporary files (optional).
            // File.Delete(sourcePath);
            // File.Delete(pdfInvariantPath);
            // File.Delete(pdfFrenchPath);
        }
    }
}