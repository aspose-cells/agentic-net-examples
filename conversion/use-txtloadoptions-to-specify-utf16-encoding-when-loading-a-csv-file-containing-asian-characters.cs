// Title: Load UTF-16 CSV with Japanese and Chinese characters using TxtLoadOptions in Aspose.Cells (.NET)
// Description: Demonstrates how to encode a CSV containing Japanese and Chinese text as UTF-16, set TxtLoadOptions.Encoding to Unicode, and import the stream into an Aspose.Cells Workbook for accurate multilingual data handling.
// Keywords: Aspose.Cells | TxtLoadOptions | UTF-16 CSV | Unicode CSV import | C# | .NET | Japanese characters | Chinese characters | multilingual CSV | Excel conversion
// Common Searches: Aspose.Cells load UTF-16 CSV | TxtLoadOptions set encoding C# | read Japanese CSV with Aspose.Cells | import Chinese characters CSV .NET | Unicode CSV to Excel Aspose
// Developer Intent: Import a CSV file that contains Asian characters by specifying UTF-16 encoding through TxtLoadOptions.
// Use Cases: Convert a UTF-16 encoded CSV with Japanese and Chinese text into an Excel workbook while preserving characters. | Read multilingual CSV data from a memory stream for further analysis or transformation in .NET applications. | Ensure correct display of non‑Latin cell values when loading CSV files into Aspose.Cells worksheets.
// AI Prompts: Show how to load a UTF-16 CSV from a file path using TxtLoadOptions in Aspose.Cells. | Provide a pattern for handling CSV files that may be UTF-8 or UTF-16, detecting BOM automatically. | Explain how to configure TxtLoadOptions to fallback to a default encoding when the file lacks a BOM.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates how to encode a CSV containing Japanese and Chinese text as UTF-16, set TxtLoadOptions.Encoding to Unicode, and import the stream into an Aspose.Cells Workbook for accurate multilingual data handling.
class LoadCsvUtf16Demo
{
    static void Main()
    {
        // CSV content that includes Asian characters (Japanese and Chinese)
        string csvContent = "名前,年齢\n山田太郎,30\n李四,25";

        // Encode the CSV data as UTF‑16 (Unicode) bytes
        byte[] utf16Bytes = Encoding.Unicode.GetBytes(csvContent);

        // Create a memory stream containing the UTF‑16 encoded CSV
        using (MemoryStream csvStream = new MemoryStream())
        {
            csvStream.Write(utf16Bytes, 0, utf16Bytes.Length);
            csvStream.Seek(0, SeekOrigin.Begin);

            // Configure TxtLoadOptions to use UTF‑16 encoding when loading the CSV
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Encoding = Encoding.Unicode   // UTF‑16
            };

            // Load the CSV into a Workbook using the specified load options
            Workbook workbook = new Workbook(csvStream, loadOptions);

            // Demonstrate that the data was loaded correctly
            Console.WriteLine("A1 (Header): " + workbook.Worksheets[0].Cells["A1"].StringValue);
            Console.WriteLine("A2 (First Name): " + workbook.Worksheets[0].Cells["A2"].StringValue);
            Console.WriteLine("B2 (First Age): " + workbook.Worksheets[0].Cells["B2"].StringValue);
        }
    }
}
