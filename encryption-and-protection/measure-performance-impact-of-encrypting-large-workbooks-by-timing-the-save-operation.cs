// Title: Benchmark Encryption Overhead When Saving a Large Excel Workbook with Aspose.Cells (.NET)
// Description: Creates a 10,000‑row by 10‑column workbook, records the time to save it as XLSX without protection, then applies a password and StrongCryptographicProvider (128‑bit) encryption and measures the save duration again, outputting both timings in milliseconds.
// Keywords: Aspose.Cells | C# encryption benchmark | Excel save performance | password protection | StrongCryptographicProvider | 128-bit encryption | large workbook | save time measurement | performance testing | Aspose.Cells .NET
// Common Searches: Aspose.Cells encryption speed test | measure save time for encrypted Excel in C# | benchmark password protection Aspose.Cells | how much time does encryption add to Excel save .NET | performance of StrongCryptographicProvider Aspose.Cells | compare encrypted vs unencrypted workbook save time
// Developer Intent: Evaluate the extra milliseconds required to write a large Excel file when password protection and strong encryption are enabled using Aspose.Cells.
// Use Cases: Validate that encrypting bulk reports meets latency requirements in a high‑throughput service. | Set performance baselines before enabling workbook protection in automated export pipelines. | Compare different encryption algorithms or key lengths to choose the optimal security‑performance trade‑off. | Document encryption overhead for service‑level agreement (SLA) planning.
// AI Prompts: Provide a C# snippet that repeats the encryption benchmark for both 128‑bit and 256‑bit keys and logs the average save times. | Show how to switch between Standard and StrongCryptographicProvider encryption in Aspose.Cells and generate a side‑by‑side timing report. | Explain how to interpret the measured milliseconds and suggest code or configuration tweaks to reduce encrypted save latency. | Create a PowerShell script that runs the compiled program, captures its output, and appends the results to a CSV file for further analysis.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsEncryptionPerformance
{
    // Creates a 10,000‑row by 10‑column workbook, records the time to save it as XLSX without protection, then applies a password and StrongCryptographicProvider (128‑bit) encryption and measures the save duration again, outputting both timings in milliseconds.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with a large amount of data (e.g., 10,000 rows × 10 columns)
            for (int row = 0; row < 10000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    // Put a simple numeric value; you can adjust the content as needed
                    sheet.Cells[row, col].PutValue(row * 10 + col);
                }
            }

            // -----------------------------------------------------------------
            // Measure save time without encryption
            // -----------------------------------------------------------------
            Stopwatch swNoEncrypt = Stopwatch.StartNew();
            workbook.Save("LargeWorkbook_NoEncryption.xlsx", SaveFormat.Xlsx);
            swNoEncrypt.Stop();
            Console.WriteLine($"Save without encryption: {swNoEncrypt.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Apply encryption settings
            // -----------------------------------------------------------------
            // Set a password for the workbook
            workbook.Settings.Password = "StrongPassword123";

            // Choose an encryption algorithm and key length (e.g., StrongCryptographicProvider with 128-bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // -----------------------------------------------------------------
            // Measure save time with encryption
            // -----------------------------------------------------------------
            Stopwatch swEncrypt = Stopwatch.StartNew();
            workbook.Save("LargeWorkbook_Encrypted.xlsx", SaveFormat.Xlsx);
            swEncrypt.Stop();
            Console.WriteLine($"Save with encryption: {swEncrypt.ElapsedMilliseconds} ms");

            // Clean up
            workbook.Dispose();
        }
    }
}
