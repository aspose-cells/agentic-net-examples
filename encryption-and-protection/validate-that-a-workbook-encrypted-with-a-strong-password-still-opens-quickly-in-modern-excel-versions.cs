// Title: Benchmark opening speed of AES‑256 encrypted Excel workbook with Aspose.Cells (.NET)
// Description: Creates a workbook, writes data to A1, applies a strong password and AES‑256 encryption, saves it, then loads the file with the password while timing the operation. The sample prints the elapsed milliseconds, verifies the IsEncrypted flag, and confirms the cell value, helping you ensure fast opening of protected files in modern Excel.
// Keywords: Aspose.Cells | C# | AES-256 encryption | password protected Excel | load performance | decryption benchmark | Excel .xlsx | EncryptionType.StrongCryptographicProvider | Stopwatch timing | IsEncrypted property
// Common Searches: Aspose.Cells measure load time encrypted Excel | C# test opening speed of password protected .xlsx | benchmark AES256 Excel decryption .NET | validate encrypted workbook opens quickly | performance test for encrypted Excel file Aspose
// Developer Intent: Validate that a workbook encrypted with a strong password opens quickly in modern Excel versions.
// Use Cases: Run automated performance tests to confirm encrypted financial reports load within acceptable thresholds. | Include decryption speed checks in CI pipelines to guarantee compliance‑driven encryption settings do not degrade user experience. | Profile encryption options when building secure Excel generators that must remain responsive for end users.
// AI Prompts: Generate C# code using Aspose.Cells that encrypts an .xlsx file with AES‑256, then measures and logs the time required to open it with the correct password. | Explain how to use the Workbook.Settings.IsEncrypted property before and after loading, and how to employ Stopwatch for accurate decryption‑time benchmarking. | Provide best‑practice recommendations for configuring strong encryption in Aspose.Cells while minimizing opening latency in .NET applications.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a workbook, writes data to A1, applies a strong password and AES‑256 encryption, saves it, then loads the file with the password while timing the operation. The sample prints the elapsed milliseconds, verifies the IsEncrypted flag, and confirms the cell value, helping you ensure fast opening of protected files in modern Excel.
class ValidateEncryptedWorkbook
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Sensitive Data");

        // Set a strong password for the workbook
        wb.Settings.Password = "StrongPassword!123";

        // Apply strong encryption (AES 256) to the workbook
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // Verify that the workbook reports being encrypted
        Console.WriteLine("Workbook IsEncrypted (original): " + wb.Settings.IsEncrypted);

        // Prepare load options with the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "StrongPassword!123";

        // Measure the time required to open the encrypted workbook
        Stopwatch sw = Stopwatch.StartNew();
        Workbook loadedWb = new Workbook(filePath, loadOptions);
        sw.Stop();

        // Output performance and verification results
        Console.WriteLine("Time to open encrypted workbook (ms): " + sw.ElapsedMilliseconds);
        Console.WriteLine("Workbook IsEncrypted (loaded): " + loadedWb.Settings.IsEncrypted);
        Console.WriteLine("Cell A1 value after load: " + loadedWb.Worksheets[0].Cells["A1"].Value);
    }
}
