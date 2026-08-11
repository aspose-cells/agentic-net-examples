// Title: Encrypt an Excel workbook with Aspose.Cells and benchmark opening time in .NET
// Description: Creates a workbook, writes sample data, applies a password with strong cryptographic encryption (128‑bit), saves the file, then loads it using LoadOptions while measuring the elapsed milliseconds with Stopwatch. The script runs headlessly and outputs the load duration and the IsEncrypted flag.
// Keywords: Aspose.Cells encrypt workbook | C# password protected Excel | measure workbook load time | headless .NET automation | LoadOptions password Aspose | encryption algorithm Aspose.Cells | performance benchmark Excel encryption
// Common Searches: How to password‑protect an Excel file using Aspose.Cells for .NET | Benchmark opening time of an encrypted workbook in C# | Supported encryption types and key lengths in Aspose.Cells | Load a password‑protected workbook with LoadOptions | Run Aspose.Cells code in a headless Docker container
// Developer Intent: Encrypt a workbook with a password and specific encryption settings, then programmatically measure how long it takes to open the protected file.
// Use Cases: Secure sensitive spreadsheets before archiving and assess performance impact in CI pipelines. | Validate encryption strength and load speed for compliance or audit requirements. | Detect regressions in encryption handling by integrating load‑time metrics into automated tests.
// AI Prompts: Generate C# code that encrypts an Excel workbook with AES‑256 using Aspose.Cells and logs the opening time in milliseconds. | Explain how to configure LoadOptions for different encryption algorithms and retrieve the IsEncrypted property after loading. | Suggest a Docker‑based, headless setup for running the opening‑time benchmark and exporting results to a monitoring system.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a workbook, writes sample data, applies a password with strong cryptographic encryption (128‑bit), saves the file, then loads it using LoadOptions while measuring the elapsed milliseconds with Stopwatch. The script runs headlessly and outputs the load duration and the IsEncrypted flag.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Encrypted data");

        // Set password to encrypt the workbook
        wb.Settings.Password = "Secret123";

        // Optional: specify encryption algorithm and key length
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        string filePath = "encrypted_workbook.xlsx";
        wb.Save(filePath);
        wb.Dispose();

        // Prepare load options with the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "Secret123";

        // Measure the time required to open the encrypted workbook
        Stopwatch sw = Stopwatch.StartNew();
        Workbook loadedWb = new Workbook(filePath, loadOptions);
        sw.Stop();

        Console.WriteLine($"Time to open encrypted workbook: {sw.ElapsedMilliseconds} ms");
        Console.WriteLine($"Workbook.IsEncrypted: {loadedWb.Settings.IsEncrypted}");

        loadedWb.Dispose();
    }
}
