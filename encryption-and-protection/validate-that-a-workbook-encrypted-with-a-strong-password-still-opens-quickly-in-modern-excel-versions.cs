// Title: How to benchmark opening speed of a strongly password‑protected XLSX workbook using Aspose.Cells in C#
// AI Prompts: Create a C# program that generates an XLSX file, applies a strong opening password with Aspose.Cells, saves it, then reloads the file using LoadOptions while measuring the elapsed milliseconds. | Write C# code that opens an existing password‑protected Excel workbook with Aspose.Cells, times the decryption process, and verifies that a specific cell contains the expected value.
// Common Searches: aspocells benchmark opening time for password protected xlsx in .net core | c# measure decryption speed of encrypted excel workbook using aspocells | how to test load performance of strong password protected workbook with aspocells | performance impact of workbook encryption in aspocells .net | load encrypted xlsx with password and get elapsed time using aspocells
// Tags: Aspose.Cells encrypted workbook load performance | C# timing password‑protected XLSX opening | LoadOptions password decryption benchmark | data integrity check after Aspose.Cells decryption | Excel opening latency measurement using Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Demonstrates creating an XLSX workbook, protecting it with a strong password via Aspose.Cells, saving it, then loading it with LoadOptions while timing the operation and confirming cell data integrity.
class WorkbookEncryptionValidation
{
    static void Main()
    {
        // Path for the workbook file
        string filePath = "EncryptedWorkbook.xlsx";

        // Strong password to encrypt the workbook
        string strongPassword = "S3cureP@ssw0rd!2026";

        try
        {
            // -------------------------
            // Create a new workbook and add sample data
            // -------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some sample data to simulate a realistic workbook
            for (int row = 0; row < 1000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // -------------------------
            // Apply encryption settings (password protection)
            // -------------------------
            workbook.Settings.Password = strongPassword; // opening password

            // Protect the workbook (structure and windows)
            workbook.Protect(ProtectionType.All, strongPassword);

            // Save the encrypted workbook
            workbook.Save(filePath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook creation/saving: {ex.Message}");
            return;
        }

        // -------------------------
        // Validate opening speed of the encrypted workbook
        // -------------------------
        Stopwatch sw = new Stopwatch();
        sw.Start();

        try
        {
            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Encrypted workbook file not found.", filePath);

            // Load the workbook with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = strongPassword
            };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            sw.Stop();

            // Output the elapsed time in milliseconds
            Console.WriteLine($"Time to open encrypted workbook: {sw.ElapsedMilliseconds} ms");

            // Simple validation: ensure the workbook opened without exception and contains expected data
            if (loadedWorkbook.Worksheets[0].Cells[0, 0].StringValue == "R0C0")
            {
                Console.WriteLine("Workbook opened successfully and data integrity verified.");
            }
            else
            {
                Console.WriteLine("Data verification failed.");
            }
        }
        catch (Exception ex)
        {
            sw.Stop();
            Console.WriteLine($"Error during workbook loading/validation: {ex.Message}");
        }
    }
}
