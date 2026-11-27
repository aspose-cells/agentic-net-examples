// Title: Export Excel cells with leading apostrophe to CSV using Aspose.Cells for .NET (preserve quote‑prefix)
// Common Searches: How can I keep the leading apostrophe when saving an Excel worksheet to CSV with Aspose.Cells in C#? | Aspose.Cells C# export quote prefix as literal character in CSV file | Save cells that have QuotePrefix enabled to CSV without losing the apostrophe | TxtSaveOptions ExportQuotePrefix true example for CSV export in .NET | Preserve apostrophe in CSV output from an Aspose.Cells workbook
// Tags: Aspose.Cells ExportQuotePrefix CSV | C# TxtSaveOptions preserve apostrophe | quote prefix to CSV conversion .NET | Excel leading apostrophe export Aspose | UTF-8 CSV output Aspose.Cells workbook

using System;
using System.IO;
using System.Text;
using System.Drawing;
using Aspose.Cells;

namespace ExportRichTextToCsv
{
    // The example creates a workbook, applies the QuotePrefix property to cells to add a leading apostrophe, and then uses TxtSaveOptions with ExportQuotePrefix set to true, UTF‑8 encoding, and a comma separator to export the active worksheet to a CSV file while preserving the apostrophe characters. The generated CSV content is read back and displayed.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for temporary files
            string excelPath = Path.Combine(Path.GetTempPath(), "RichTextDemo.xlsx");
            string csvPath = Path.Combine(Path.GetTempPath(), "RichTextExport.csv");

            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook and add content with an apostrophe
                // ------------------------------------------------------------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Cell A1: plain text with a leading apostrophe (quote prefix)
                Cell cellA1 = cells["A1"];
                cellA1.PutValue("Sample");
                Style styleA1 = cellA1.GetStyle();
                styleA1.QuotePrefix = true;               // enable quote prefix
                cellA1.SetStyle(styleA1);

                // Cell B1: simple text (rich‑text not required for CSV export)
                Cell cellB1 = cells["B1"];
                cellB1.PutValue("HelloWorld");
                Style styleB1 = cellB1.GetStyle();
                styleB1.QuotePrefix = true;               // enable quote prefix
                cellB1.SetStyle(styleB1);

                // ------------------------------------------------------------
                // 2. Save the workbook as an Excel file (optional, just for inspection)
                // ------------------------------------------------------------
                wb.Save(excelPath, SaveFormat.Xlsx);
                Console.WriteLine($"Excel workbook saved to: {excelPath}");

                // ------------------------------------------------------------
                // 3. Configure TxtSaveOptions to export apostrophes literally
                // ------------------------------------------------------------
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = ',',               // standard CSV separator
                    Encoding = Encoding.UTF8,      // UTF‑8 encoding
                    ExportAllSheets = false,       // only active sheet
                    ExportQuotePrefix = true       // export the leading apostrophe
                };

                // ------------------------------------------------------------
                // 4. Export the worksheet to CSV
                // ------------------------------------------------------------
                wb.Save(csvPath, saveOptions);
                Console.WriteLine($"CSV file exported to: {csvPath}");

                // ------------------------------------------------------------
                // 5. Verify the CSV content – read and display the lines
                // ------------------------------------------------------------
                if (File.Exists(csvPath))
                {
                    Console.WriteLine("\n--- CSV Content ---");
                    foreach (string line in File.ReadAllLines(csvPath, Encoding.UTF8))
                    {
                        Console.WriteLine(line);
                    }
                }
                else
                {
                    Console.WriteLine("CSV file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Clean up temporary files if desired
                // if (File.Exists(excelPath)) File.Delete(excelPath);
                // if (File.Exists(csvPath)) File.Delete(csvPath);
            }
        }
    }
}
