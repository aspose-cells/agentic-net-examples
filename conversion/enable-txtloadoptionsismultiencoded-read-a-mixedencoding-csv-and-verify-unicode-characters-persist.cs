// Title: C# – Load Mixed‑Encoding CSV (UTF‑8 & UTF‑16LE) with TxtLoadOptions.IsMultiEncoded in Aspose.Cells
// Description: Demonstrates how to create a temporary CSV that combines UTF‑8, UTF‑8 Cyrillic, and UTF‑16LE Japanese sections, load it with Aspose.Cells using TxtLoadOptions.IsMultiEncoded, verify that Unicode characters are preserved, and optionally save the result as an XLSX workbook.
// Keywords: Aspose.Cells | TxtLoadOptions | IsMultiEncoded | mixed encoding CSV | UTF-8 UTF-16LE | Unicode import .NET | C# CSV to Excel | multi‑encoding detection | preserve Unicode characters
// Common Searches: Aspose.Cells read CSV with multiple encodings | IsMultiEncoded true example C# | load UTF‑16LE data inside UTF‑8 CSV Aspose.Cells | verify Unicode characters after CSV import .NET | convert mixed‑encoding CSV to XLSX
// Developer Intent: Enable Aspose.Cells to detect and load sections of a CSV file encoded in different Unicode formats, ensuring all characters remain intact after import.
// Use Cases: Import legacy CSV exports where each segment uses a distinct encoding and convert them to Excel for analysis. | Validate that Cyrillic and Japanese text are correctly imported from international data feeds. | Create an automated pipeline that reads mixed‑encoding CSV files and saves them as XLSX while preserving original character sets.
// AI Prompts: Write C# code that uses Aspose.Cells TxtLoadOptions with IsMultiEncoded set to true to read a CSV containing UTF‑8 and UTF‑16LE sections and save it as an XLSX file. | Explain how Aspose.Cells determines encoding boundaries when IsMultiEncoded is enabled and which fallback encoding is applied. | Provide a step‑by‑step tutorial for generating a mixed‑encoding CSV in memory, loading it with Aspose.Cells, and checking specific cells for expected Unicode strings.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMixedEncodingDemo
{
    // Demonstrates how to create a temporary CSV that combines UTF‑8, UTF‑8 Cyrillic, and UTF‑16LE Japanese sections, load it with Aspose.Cells using TxtLoadOptions.IsMultiEncoded, verify that Unicode characters are preserved, and optionally save the result as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // Prepare mixed‑encoding CSV content
            // Part 1 – UTF‑8 encoded (default)
            string part1 = "Name,Comment\nJohn,Hello\n";

            // Part 2 – UTF‑8 encoded Unicode (Cyrillic)
            string part2 = "Anna,Привет\n";

            // Part 3 – UTF‑16LE encoded Unicode (Japanese)
            string part3 = "Bob,こんにちは\n";

            // Convert each part to bytes with the appropriate encoding
            byte[] bytesPart1 = Encoding.UTF8.GetBytes(part1);
            byte[] bytesPart2 = Encoding.UTF8.GetBytes(part2);
            byte[] bytesPart3 = Encoding.Unicode.GetBytes(part3); // UTF‑16LE

            // Combine the byte arrays into a single CSV file
            string tempCsvPath = Path.GetTempFileName();
            using (FileStream fs = new FileStream(tempCsvPath, FileMode.Create, FileAccess.Write))
            {
                fs.Write(bytesPart1, 0, bytesPart1.Length);
                fs.Write(bytesPart2, 0, bytesPart2.Length);
                fs.Write(bytesPart3, 0, bytesPart3.Length);
            }

            // Configure TxtLoadOptions to handle multiple encodings
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Separator = ',',          // CSV separator
                IsMultiEncoded = true,    // Enable multi‑encoding detection
                Encoding = Encoding.UTF8   // Default encoding for sections without BOM
            };

            // Load the mixed‑encoding CSV into a workbook using the options
            Workbook workbook = new Workbook(tempCsvPath, loadOptions);

            // Access the loaded cells to verify Unicode characters are preserved
            Cells cells = workbook.Worksheets[0].Cells;

            Console.WriteLine("A2 (John): " + cells["A2"].StringValue);          // John
            Console.WriteLine("B2 (Hello): " + cells["B2"].StringValue);        // Hello

            Console.WriteLine("A3 (Anna): " + cells["A3"].StringValue);          // Anna
            Console.WriteLine("B3 (Cyrillic): " + cells["B3"].StringValue);     // Привет

            Console.WriteLine("A4 (Bob): " + cells["A4"].StringValue);           // Bob
            Console.WriteLine("B4 (Japanese): " + cells["B4"].StringValue);     // こんにちは

            // Optional: save the workbook to verify successful load (uses default save logic)
            string outputPath = Path.Combine(Path.GetDirectoryName(tempCsvPath), "MixedEncodingOutput.xlsx");
            workbook.Save(outputPath);

            // Clean up temporary CSV file
            File.Delete(tempCsvPath);
        }
    }
}
