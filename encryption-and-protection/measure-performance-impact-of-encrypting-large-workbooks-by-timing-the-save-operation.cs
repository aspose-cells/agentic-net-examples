// Title: Aspose.Cells .NET Benchmark: Save Speed of Large Workbook With and Without Encryption
// Description: Creates a 50,000‑row by 10‑column Excel workbook, records the time to save it as an unencrypted XLSX file, then applies a password and 128‑bit StrongCryptographicProvider encryption via Workbook.Settings.Password and Workbook.SetEncryptionOptions, saves the encrypted file, and prints both durations. Use this sample to evaluate the performance impact of encryption on large Excel exports with Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | encryption performance | save benchmark | large workbook | StrongCryptographicProvider | 128-bit encryption | Workbook.Settings.Password | SetEncryptionOptions | Excel export speed | measure encryption overhead
// Common Searches: Aspose.Cells encryption performance benchmark | how long does saving an encrypted workbook take .NET | benchmark Aspose.Cells save with password | measure save time for large Excel file Aspose.Cells | C# Aspose.Cells encrypt workbook speed
// Developer Intent: The developer wants to measure and compare the time required to save a large Excel workbook with and without password‑based encryption using Aspose.Cells for .NET.
// Use Cases: Determine whether encryption is acceptable for time‑critical large report generation. | Set performance thresholds for enabling or disabling encryption in automated export pipelines. | Compare encryption overhead across different workbook sizes or key lengths. | Integrate save‑time metrics into CI/CD tests to catch regressions. | Provide data for capacity planning of server‑side Excel processing.
// AI Prompts: Generate a C# example that benchmarks Aspose.Cells save time using AES‑256 encryption and reports the results. | Explain how to interpret the timing output and suggest ways to reduce encryption overhead in Aspose.Cells. | Create a unit test that asserts the encrypted save operation for a 50,000‑row workbook completes within a specified time limit. | Provide a PowerShell script that runs the benchmark on multiple machines and aggregates the results.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsEncryptionPerformance
{
    // Creates a 50,000‑row by 10‑column Excel workbook, records the time to save it as an unencrypted XLSX file, then applies a password and 128‑bit StrongCryptographicProvider encryption via Workbook.Settings.Password and Workbook.SetEncryptionOptions, saves the encrypted file, and prints both durations. Use this sample to evaluate the performance impact of encryption on large Excel exports with Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with a large amount of data
            // Example: 50,000 rows x 10 columns
            const int totalRows = 50000;
            const int totalCols = 10;
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Fill each cell with a simple string value
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ------------------------------
            // Measure save time without encryption
            // ------------------------------
            Stopwatch swNoEncrypt = Stopwatch.StartNew();
            // Save using the Save(string, SaveFormat) rule
            workbook.Save("LargeWorkbook_NoEncryption.xlsx", SaveFormat.Xlsx);
            swNoEncrypt.Stop();
            Console.WriteLine($"Save without encryption: {swNoEncrypt.ElapsedMilliseconds} ms");

            // ------------------------------
            // Apply encryption settings
            // ------------------------------
            // Set a password for the workbook (WorkbookSettings.Password property rule)
            workbook.Settings.Password = "StrongPassword123";

            // Set encryption options (Workbook.SetEncryptionOptions method rule)
            // Using StrongCryptographicProvider with a 128‑bit key
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // ------------------------------
            // Measure save time with encryption
            // ------------------------------
            Stopwatch swEncrypt = Stopwatch.StartNew();
            // Save the encrypted workbook (same Save rule)
            workbook.Save("LargeWorkbook_Encrypted.xlsx", SaveFormat.Xlsx);
            swEncrypt.Stop();
            Console.WriteLine($"Save with encryption: {swEncrypt.ElapsedMilliseconds} ms");

            // Clean up
            workbook.Dispose();
        }
    }
}
