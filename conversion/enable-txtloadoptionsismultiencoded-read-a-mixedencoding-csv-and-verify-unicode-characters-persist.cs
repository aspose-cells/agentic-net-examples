// Title: Import a Mixed‑Encoding CSV with TxtLoadOptions.IsMultiEncoded in Aspose.Cells for .NET
// Description: This example builds a MemoryStream that concatenates UTF‑8 and UTF‑16LE byte sequences, activates TxtLoadOptions.IsMultiEncoded, sets the comma separator, and loads the stream into a worksheet via ImportCSV. After import, the code prints cell values to confirm that Japanese characters are intact and optionally saves the workbook as an XLSX file.
// Keywords: Aspose.Cells | TxtLoadOptions.IsMultiEncoded | ImportCSV | C# mixed encoding CSV | UTF-8 UTF-16LE | preserve Unicode | memory stream CSV | multi‑encoded CSV | Excel export
// Common Searches: Aspose.Cells read CSV with multiple encodings | Enable IsMultiEncoded for CSV import | Load UTF‑8 and UTF‑16LE data in one CSV | Check Unicode characters after ImportCSV | C# Aspose.Cells mixed encoding example
// Developer Intent: Load a CSV that contains sections encoded in different character sets using TxtLoadOptions.IsMultiEncoded and verify that non‑ASCII text remains unchanged.
// Use Cases: Processing logs that mix English and localized strings | Migrating legacy CSV files with mixed encodings to XLSX | Automated validation of Unicode data during import | Reading CSV streams from APIs that may switch encodings | Generating reports from multilingual CSV sources
// AI Prompts: Generate C# code that reads a CSV stream containing UTF‑8 and UTF‑16LE data with Aspose.Cells TxtLoadOptions.IsMultiEncoded and writes the result to an XLSX file. | Explain how TxtLoadOptions.IsMultiEncoded detects encoding changes within a single CSV stream and list any known limitations. | Create a C# unit test that asserts Japanese characters are preserved after importing a mixed‑encoding CSV using Aspose.Cells. | Suggest performance‑friendly ways to import large multi‑encoded CSV files with Aspose.Cells. | Troubleshoot why Unicode characters might appear as garbled text when IsMultiEncoded is set to false.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMixedEncodingDemo
{
    // This example builds a MemoryStream that concatenates UTF‑8 and UTF‑16LE byte sequences, activates TxtLoadOptions.IsMultiEncoded, sets the comma separator, and loads the stream into a worksheet via ImportCSV. After import, the code prints cell values to confirm that Japanese characters are intact and optionally saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // ----- Prepare mixed‑encoding CSV data -----
            // Part 1: UTF‑8 encoded (ASCII characters)
            string partUtf8 = "Name,Comment\nJohn,Hello\n";

            // Part 2: UTF‑16LE encoded (Japanese characters)
            string partUtf16 = "Anna,こんにちは\n";

            // Convert each part to its respective byte representation
            byte[] bytesUtf8 = Encoding.UTF8.GetBytes(partUtf8);
            byte[] bytesUtf16 = Encoding.Unicode.GetBytes(partUtf16); // Unicode = UTF‑16LE

            // Combine the two byte arrays into a single stream
            MemoryStream mixedStream = new MemoryStream();
            mixedStream.Write(bytesUtf8, 0, bytesUtf8.Length);
            mixedStream.Write(bytesUtf16, 0, bytesUtf16.Length);
            mixedStream.Position = 0; // Reset for reading

            // ----- Configure TxtLoadOptions -----
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.IsMultiEncoded = true;      // Enable handling of multiple encodings
            loadOptions.Separator = ',';            // CSV separator
            loadOptions.ConvertNumericData = false; // Keep all data as strings for this demo

            // ----- Load CSV into a workbook using ImportCSV (stream overload) -----
            Workbook workbook = new Workbook(); // Empty workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.ImportCSV(mixedStream, loadOptions, 0, 0);

            // ----- Verify that Unicode characters are preserved -----
            // Row 1 (index 0) contains header, Row 2 (index 1) = John, Row 3 (index 2) = Anna
            Console.WriteLine("A2 (Name): " + sheet.Cells["A2"].StringValue); // Expected: John
            Console.WriteLine("B2 (Comment): " + sheet.Cells["B2"].StringValue); // Expected: Hello
            Console.WriteLine("A3 (Name): " + sheet.Cells["A3"].StringValue); // Expected: Anna
            Console.WriteLine("B3 (Comment): " + sheet.Cells["B3"].StringValue); // Expected: こんにちは

            // Optional: Save to verify visually (uses default UTF‑8 encoding)
            workbook.Save("MixedEncodingOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}
