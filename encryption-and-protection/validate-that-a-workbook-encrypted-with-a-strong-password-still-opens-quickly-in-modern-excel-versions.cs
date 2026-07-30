// Title: C# performance test: opening an AES‑256 encrypted Excel workbook with Aspose.Cells
// Description: This example creates a workbook, writes sample data, applies a strong password and AES‑256 encryption using Aspose.Cells for .NET, saves it as .xlsx, then loads it with LoadOptions while timing the operation with Stopwatch. It outputs the elapsed milliseconds and verifies the encryption flag before and after loading.
// Keywords: Aspose.Cells | C# encryption | AES-256 Excel | password protected workbook | load performance | benchmark Excel decryption | strong password | SaveFormat.Xlsx | LoadOptions | Stopwatch timing
// Common Searches: How fast can Aspose.Cells open an AES‑256 encrypted .xlsx in C#? | Measure load time of password‑protected Excel file using Aspose.Cells | C# code to benchmark encrypted workbook opening | Performance of strong encryption in Aspose.Cells .NET | Open encrypted Excel workbook quickly
// Developer Intent: Confirm that a workbook encrypted with a strong password and AES‑256 can be opened rapidly with Aspose.Cells, matching modern Excel performance expectations.
// Use Cases: Benchmark decryption latency for compliance‑driven Excel files | Validate that encryption does not degrade user experience in enterprise applications | Demonstrate setting encryption options and measuring load time programmatically | Ensure encrypted workbooks remain flagged as encrypted after save and load
// AI Prompts: Write C# code using Aspose.Cells to encrypt an .xlsx with AES‑256 and measure the time to open it with a password. | Provide a unit‑test snippet that asserts the opening time of an encrypted workbook is under a given threshold. | Explain how to configure LoadOptions for a password‑protected Excel file and check the IsEncrypted property after loading. | Suggest ways to log performance metrics for encrypted workbook loading in a .NET application.

using System;
using System.Diagnostics;
using Aspose.Cells;

// This example creates a workbook, writes sample data, applies a strong password and AES‑256 encryption using Aspose.Cells for .NET, saves it as .xlsx, then loads it with LoadOptions while timing the operation with Stopwatch. It outputs the elapsed milliseconds and verifies the encryption flag before and after loading.
class WorkbookEncryptionPerformanceDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Performance test data");

        // Set a strong password for encryption
        wb.Settings.Password = "Str0ngP@ssw0rd!2026";

        // Apply strong encryption (AES 256)
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        string filePath = "EncryptedPerformance.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // Verify that the workbook is marked as encrypted
        Console.WriteLine("IsEncrypted after save: " + wb.Settings.IsEncrypted);

        // Prepare load options with the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "Str0ngP@ssw0rd!2026";

        // Measure the time required to open the encrypted workbook
        Stopwatch sw = Stopwatch.StartNew();
        Workbook loadedWb = new Workbook(filePath, loadOptions);
        sw.Stop();

        Console.WriteLine($"Time to open encrypted workbook: {sw.ElapsedMilliseconds} ms");
        Console.WriteLine("IsEncrypted after load: " + loadedWb.Settings.IsEncrypted);
    }
}
