// Title: Benchmark Encryption Overhead When Saving Large Excel Workbooks with Aspose.Cells for .NET
// Description: A C# example that builds a 10,000‑row × 10‑column workbook, records the time to save it as an unprotected XLSX file, then applies a password and StrongCryptographicProvider (128‑bit) encryption, saves again, and outputs both timings. The data helps evaluate the performance cost of encryption in Aspose.Cells for .NET.
// Keywords: Aspose.Cells encryption benchmark | C# workbook save performance | large Excel file save time | password protection Aspose.Cells | .NET Excel encryption overhead | StrongCryptographicProvider 128‑bit | measure save duration Aspose | Excel file encryption timing | performance testing Aspose.Cells | Stopwatch save time C#
// Common Searches: Aspose.Cells how long does encrypted save take | benchmark saving large workbook without password .NET | encryption overhead Aspose.Cells StrongCryptographicProvider | measure Excel file save time with password C# | performance impact of workbook encryption Aspose
// Developer Intent: Determine the time penalty introduced by password‑based encryption when saving a large workbook with Aspose.Cells for .NET.
// Use Cases: Establish a baseline save speed for a massive workbook before applying any protection. | Quantify the additional milliseconds caused by StrongCryptographicProvider encryption. | Compare unencrypted vs encrypted timings to decide if encryption meets performance requirements. | Provide data for capacity planning in high‑throughput Excel generation pipelines.
// AI Prompts: Write C# code that logs both unencrypted and encrypted save times using Aspose.Cells and calculates the percentage slowdown. | Explain how to switch between EncryptionType.Standard and EncryptionType.StrongCryptographicProvider in Aspose.Cells and benchmark each option on a large workbook. | Give guidance on interpreting the measured timings to choose suitable encryption settings for a high‑volume Excel export service.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // A C# example that builds a 10,000‑row × 10‑column workbook, records the time to save it as an unprotected XLSX file, then applies a password and StrongCryptographicProvider (128‑bit) encryption, saves again, and outputs both timings. The data helps evaluate the performance cost of encryption in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (rule: Workbook())
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with a large amount of data
            // Example: 10,000 rows x 10 columns
            for (int row = 0; row < 10000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    // Put a simple value into each cell
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // ------------------------------
            // Measure save time without encryption
            // ------------------------------
            Stopwatch swNoEncrypt = Stopwatch.StartNew();

            // Save the workbook without any encryption (rule: Save(string, SaveFormat))
            workbook.Save("LargeWorkbook_NoEncryption.xlsx", SaveFormat.Xlsx);

            swNoEncrypt.Stop();
            Console.WriteLine($"Save without encryption elapsed: {swNoEncrypt.ElapsedMilliseconds} ms");

            // ------------------------------
            // Apply encryption settings
            // ------------------------------

            // Set a password for the workbook (rule: WorkbookSettings.Password)
            workbook.Settings.Password = "StrongPassword123";

            // Set encryption options (rule: SetEncryptionOptions)
            // Using StrongCryptographicProvider with 128-bit key
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // ------------------------------
            // Measure save time with encryption
            // ------------------------------
            Stopwatch swEncrypt = Stopwatch.StartNew();

            // Save the encrypted workbook (rule: Save(string, SaveFormat))
            workbook.Save("LargeWorkbook_Encrypted.xlsx", SaveFormat.Xlsx);

            swEncrypt.Stop();
            Console.WriteLine($"Save with encryption elapsed: {swEncrypt.ElapsedMilliseconds} ms");

            // Clean up
            workbook.Dispose();
        }
    }
}
