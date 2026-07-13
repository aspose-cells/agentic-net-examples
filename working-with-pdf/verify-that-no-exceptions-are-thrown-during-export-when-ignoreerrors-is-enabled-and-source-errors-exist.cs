using System;
using Aspose.Cells;
using System.Data;

namespace AsposeCellsIgnoreErrorDemo
{
    // Author: Aspose.Cells .NET example demonstrating IgnoreError during export
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a formula that will cause a calculation error (non‑existent function)
            sheet.Cells["A1"].Formula = "=NONEXISTENTFUNC()";

            // ------------------------------
            // 1. Verify export with PaginatedSaveOptions.IgnoreError
            // ------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Enable ignoring rendering errors (e.g., shape, chart, or formula errors)
                IgnoreError = true
            };

            try
            {
                // Save as PDF; no exception should be thrown despite the formula error
                workbook.Save("ExportWithIgnoreError.pdf", pdfOptions);
                Console.WriteLine("PDF saved successfully with IgnoreError enabled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected exception during PDF export: " + ex.Message);
            }

            // ------------------------------
            // 2. Verify export to DataTable with ExportTableOptions.SkipErrorValue
            // ------------------------------
            // Place a value that exceeds decimal.MaxValue to trigger a conversion error
            sheet.Cells["B1"].Value = decimal.MaxValue;
            sheet.Cells["B2"].Value = (double)decimal.MaxValue + 1e10; // invalid for decimal column

            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportColumnName = true,
                SkipErrorValue = true,      // Instruct exporter to skip invalid values
                CheckMixedValueType = true  // Enable mixed type checking
            };

            try
            {
                // Export the range (A1:B2) to a DataTable; invalid value should be skipped without exception
                DataTable dt = sheet.Cells.ExportDataTable(0, 0, 3, 2, exportOptions);
                Console.WriteLine("DataTable exported successfully with SkipErrorValue enabled.");
                Console.WriteLine("Rows exported: " + dt.Rows.Count);
                // Display exported values for verification
                foreach (DataRow row in dt.Rows)
                {
                    Console.WriteLine(string.Join(", ", row.ItemArray));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected exception during DataTable export: " + ex.Message);
            }
        }
    }
}