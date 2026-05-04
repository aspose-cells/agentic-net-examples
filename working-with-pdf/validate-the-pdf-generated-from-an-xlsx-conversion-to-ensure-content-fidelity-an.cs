using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Utility;

namespace AsposeCellsPdfValidation
{
    class Program
    {
        static void Main()
        {
            // Paths for source Excel and generated PDF
            string sourceExcelPath = "sample.xlsx";
            string generatedPdfPath = "sample_converted.pdf";

            // -------------------------------------------------
            // 1. Load the Excel workbook (create/load lifecycle)
            // -------------------------------------------------
            // LoadOptions can be customized if needed; using default here
            Workbook workbook = new Workbook(sourceExcelPath);

            // -------------------------------------------------
            // 2. Configure PDF save options
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure PDF/A-1b compliance for better interoperability
                Compliance = PdfCompliance.PdfA1b,

                // Enable font compatibility checking to avoid missing glyphs
                CheckFontCompatibility = true,

                // Use workbook's default font first for Unicode characters
                CheckWorkbookDefaultFont = true,

                // Do not ignore rendering errors; we want them to surface
                IgnoreError = false,

                // Calculate formulas before rendering so results appear in PDF
                CalculateFormula = true
            };

            // -------------------------------------------------
            // 3. Convert Excel to PDF using the configured options
            // -------------------------------------------------
            // Using the Workbook.Save method with PdfSaveOptions (lifecycle rule)
            workbook.Save(generatedPdfPath, pdfOptions);

            // -------------------------------------------------
            // 4. Basic validation of the generated PDF
            // -------------------------------------------------
            // Verify that the PDF file was created
            if (!File.Exists(generatedPdfPath))
            {
                Console.WriteLine("Error: PDF file was not created.");
                return;
            }

            // Verify that the PDF file size is reasonable (greater than 1 KB)
            FileInfo pdfInfo = new FileInfo(generatedPdfPath);
            if (pdfInfo.Length < 1024)
            {
                Console.WriteLine($"Warning: PDF file size is unusually small ({pdfInfo.Length} bytes).");
            }
            else
            {
                Console.WriteLine($"PDF generated successfully. Size: {pdfInfo.Length} bytes.");
            }

            // Optional: Re‑convert the PDF back to Excel (round‑trip) and compare cell values
            // This demonstrates layout fidelity by ensuring data round‑trips correctly.
            string roundTripExcelPath = "roundtrip.xlsx";

            // Use ConversionUtility to convert PDF back to XLSX (if supported)
            // Note: ConversionUtility can handle many formats; here we attempt PDF->XLSX.
            try
            {
                ConversionUtility.Convert(generatedPdfPath, roundTripExcelPath);
                Console.WriteLine("Round‑trip conversion (PDF → XLSX) completed.");

                // Load the round‑trip workbook and compare a few cells
                Workbook roundTripWorkbook = new Workbook(roundTripExcelPath);
                Worksheet originalSheet = workbook.Worksheets[0];
                Worksheet roundTripSheet = roundTripWorkbook.Worksheets[0];

                // Simple cell comparison for the first few rows/columns
                bool dataMatches = true;
                for (int row = 0; row < Math.Min(10, originalSheet.Cells.MaxDataRow + 1); row++)
                {
                    for (int col = 0; col < Math.Min(5, originalSheet.Cells.MaxDataColumn + 1); col++)
                    {
                        var originalValue = originalSheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                        var roundTripValue = roundTripSheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                        if (!originalValue.Equals(roundTripValue))
                        {
                            dataMatches = false;
                            Console.WriteLine($"Mismatch at cell {CellIndexToName(row, col)}: " +
                                              $"original='{originalValue}' vs round‑trip='{roundTripValue}'");
                        }
                    }
                }

                Console.WriteLine(dataMatches
                    ? "Data fidelity check passed: cell values match after round‑trip."
                    : "Data fidelity check failed: some cell values differ after round‑trip.");
            }
            catch (Exception ex)
            {
                // If PDF→XLSX conversion is not supported, log the exception but do not treat as fatal
                Console.WriteLine($"Round‑trip conversion not performed: {ex.Message}");
            }
        }

        // Helper to convert zero‑based row/column indices to Excel cell name (e.g., 0,0 -> A1)
        private static string CellIndexToName(int row, int column)
        {
            int dividend = column + 1;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return $"{columnName}{row + 1}";
        }
    }
}