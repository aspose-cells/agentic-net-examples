// Title: Measure load time of a password‑protected Excel file using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, applies a password with strong 128‑bit encryption, saves it, then opens the file with LoadOptions while timing the operation via Stopwatch and verifies the workbook remains encrypted.
// Keywords: Aspose.Cells | .NET | encrypt workbook | password protected Excel | load time measurement | Stopwatch | LoadOptions | encryption algorithm | 128-bit encryption | performance benchmark
// Common Searches: Aspose.Cells encrypt Excel and measure opening speed | C# benchmark loading encrypted workbook | How to time loading of password protected Excel with Aspose.Cells | LoadOptions password example Aspose.Cells | Performance impact of Excel encryption .NET
// Developer Intent: Encrypt an Excel workbook with a password and defined encryption settings, then programmatically determine how long it takes to open the encrypted file.
// Use Cases: Compare load performance of different encryption algorithms or key lengths for automated reporting pipelines. | Validate that encrypted workbooks meet latency requirements in headless CI/CD environments. | Run security compliance tests that ensure encrypted files can be opened reliably within expected time frames.
// AI Prompts: Generate C# code that encrypts an Aspose.Cells workbook with AES‑256, saves it, and logs the opening time using LoadOptions and Stopwatch. | Show how to benchmark opening times for workbooks encrypted with 128‑bit versus 256‑bit keys in Aspose.Cells. | Explain how to handle incorrect passwords and capture load‑time metrics for encrypted Excel files in a headless automation scenario.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsEncryptionTiming
{
    // Creates a workbook, adds sample data, applies a password with strong 128‑bit encryption, saves it, then opens the file with LoadOptions while timing the operation via Stopwatch and verifies the workbook remains encrypted.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Test");
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Set a password to encrypt the workbook
            workbook.Settings.Password = "SecretPassword123";

            // Optional: specify encryption algorithm and key length
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook to disk (uses Workbook.Save(string))
            string encryptedFilePath = "encrypted.xlsx";
            workbook.Save(encryptedFilePath);

            // Prepare load options with the same password for opening
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "SecretPassword123"
            };

            // Measure the time required to open the encrypted workbook
            Stopwatch stopwatch = Stopwatch.StartNew();

            // Load the workbook using the constructor that accepts file path and LoadOptions
            Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);

            stopwatch.Stop();

            // Output the elapsed time in milliseconds
            Console.WriteLine($"Time to open encrypted workbook: {stopwatch.ElapsedMilliseconds} ms");

            // Verify that the workbook is indeed encrypted
            Console.WriteLine($"IsEncrypted property after load: {loadedWorkbook.Settings.IsEncrypted}");

            // Clean up
            workbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}
