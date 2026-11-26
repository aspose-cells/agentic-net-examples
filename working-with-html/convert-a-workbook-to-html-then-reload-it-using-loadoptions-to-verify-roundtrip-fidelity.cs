// Title: Save an Excel workbook as HTML and reload it with LoadOptions to verify round‑trip fidelity using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook to HTML and load it back for verification | how to use LoadOptions with LoadFormat.Html to reload HTML in Aspose.Cells | compare original Excel workbook with workbook loaded from HTML using Aspose.Cells | C# round‑trip conversion from XLSX to HTML and back with Aspose.Cells | verify cell values after converting XLSX to HTML with Aspose.Cells
// Tags: HtmlSaveOptions export gridlines base64 images | LoadOptions LoadFormat.Html workbook reload | round‑trip fidelity HTML conversion Aspose.Cells | cell value comparison after HTML conversion C# | sample XLSX workbook creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlRoundTrip
{
    // The program creates a sample XLSX workbook if missing, saves it as HTML using HtmlSaveOptions (with gridlines and base64‑encoded images), reloads the HTML into a new Workbook via LoadOptions(LoadFormat.Html), and compares worksheet counts and cell values to confirm round‑trip fidelity.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define file paths
                string workbookPath = "sample.xlsx";
                string htmlPath = "sample.html";

                // Ensure a source workbook exists
                if (!File.Exists(workbookPath))
                {
                    CreateSampleWorkbook(workbookPath);
                    Console.WriteLine($"Created sample workbook at '{workbookPath}'.");
                }

                // Load the original workbook
                Workbook originalWorkbook = new Workbook(workbookPath);
                Console.WriteLine("Original workbook loaded.");

                // Convert workbook to HTML
                HtmlSaveOptions htmlSaveOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,
                    ExportImagesAsBase64 = true
                    // SaveFormat is read‑only for HtmlSaveOptions; no need to set it explicitly
                };

                originalWorkbook.Save(htmlPath, htmlSaveOptions);
                Console.WriteLine($"Workbook saved as HTML to '{htmlPath}'.");

                // Verify that the HTML file was created before loading
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: HTML file '{htmlPath}' was not created.");
                    return;
                }

                // Load the HTML back into a new workbook using LoadOptions
                Workbook htmlLoadedWorkbook;
                try
                {
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                    htmlLoadedWorkbook = new Workbook(htmlPath, loadOptions);
                    Console.WriteLine("HTML file loaded back into a workbook.");
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine("Failed to load HTML file:");
                    Console.WriteLine(loadEx.Message);
                    return;
                }

                // Verify round‑trip fidelity
                bool isIdentical = CompareWorkbooks(originalWorkbook, htmlLoadedWorkbook);
                Console.WriteLine($"Round‑trip fidelity verification result: {(isIdentical ? "PASS" : "FAIL")}");

                // Optional: output differences if any
                if (!isIdentical)
                {
                    Console.WriteLine("Differences were detected between the original and reloaded workbooks.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="path">File path where the workbook will be saved.</param>
        private static void CreateSampleWorkbook(string path)
        {
            try
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Populate some sample data
                ws.Cells["A1"].PutValue("Product");
                ws.Cells["B1"].PutValue("Price");
                ws.Cells["A2"].PutValue("Apple");
                ws.Cells["B2"].PutValue(1.25);
                ws.Cells["A3"].PutValue("Banana");
                ws.Cells["B3"].PutValue(0.80);
                ws.Cells["A4"].PutValue("Total");
                ws.Cells["B4"].Formula = "=SUM(B2:B3)";

                // Calculate formulas so that values are stored
                wb.CalculateFormula();

                // Save the workbook
                wb.Save(path, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Compares two workbooks for basic fidelity: worksheet count and cell values in the first worksheet.
        /// </summary>
        /// <param name="wb1">Original workbook.</param>
        /// <param name="wb2">Workbook loaded from HTML.</param>
        /// <returns>True if the compared aspects are identical; otherwise false.</returns>
        private static bool CompareWorkbooks(Workbook wb1, Workbook wb2)
        {
            // Compare worksheet count
            if (wb1.Worksheets.Count != wb2.Worksheets.Count)
            {
                Console.WriteLine($"Worksheet count mismatch: {wb1.Worksheets.Count} vs {wb2.Worksheets.Count}");
                return false;
            }

            // Compare cell values in the first worksheet (up to a reasonable range)
            Worksheet ws1 = wb1.Worksheets[0];
            Worksheet ws2 = wb2.Worksheets[0];
            Cells cells1 = ws1.Cells;
            Cells cells2 = ws2.Cells;

            // Determine the used range to compare
            int maxRow = Math.Max(cells1.MaxDataRow, cells2.MaxDataRow);
            int maxCol = Math.Max(cells1.MaxDataColumn, cells2.MaxDataColumn);

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    string val1 = cells1[row, col].StringValue;
                    string val2 = cells2[row, col].StringValue;

                    if (!string.Equals(val1, val2, StringComparison.Ordinal))
                    {
                        Console.WriteLine($"Cell mismatch at {CellIndexToName(row, col)}: \"{val1}\" vs \"{val2}\"");
                        return false;
                    }
                }
            }

            // If we reach here, basic checks passed
            return true;
        }

        /// <summary>
        /// Converts zero‑based row/column indices to an Excel cell name (e.g., 0,0 => A1).
        /// </summary>
        private static string CellIndexToName(int row, int column)
        {
            // Convert column index to letters
            string colName = "";
            int dividend = column + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colName = Convert.ToChar('A' + modulo) + colName;
                dividend = (dividend - modulo) / 26;
            }

            // Row numbers are 1‑based
            return $"{colName}{row + 1}";
        }
    }
}
