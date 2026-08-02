// Title: Benchmark Opening Time of an Encrypted Excel Workbook with Aspose.Cells for .NET
// Description: C# sample that loads a password‑protected Excel file using Aspose.Cells LoadOptions, measures the elapsed time with Stopwatch, prints the duration in milliseconds, displays the workbook's encryption flags, and releases resources.
// Keywords: Aspose.Cells encrypted workbook performance | measure load time encrypted Excel .NET | benchmark Aspose.Cells workbook opening | Stopwatch load duration password protected Excel | IsEncrypted Aspose.Cells | IsWorkbookProtectedWithPassword | Excel decryption timing .NET | performance metrics Aspose.Cells
// Common Searches: how to benchmark opening an encrypted Excel file with Aspose.Cells C# | measure load time of password protected workbook Aspose.Cells .NET | log encryption status after loading encrypted workbook Aspose.Cells | Aspose.Cells performance test for encrypted workbooks | C# code to time opening of protected Excel file using Aspose
// Developer Intent: Determine the duration required to load a password‑protected workbook and capture its encryption status.
// Use Cases: Run automated performance tests on large encrypted workbooks. | Detect regressions in decryption speed after library upgrades. | Record encryption flags for compliance reporting in CI pipelines.
// AI Prompts: Generate C# code that times the opening of an encrypted Excel workbook with Aspose.Cells and outputs both elapsed milliseconds and encryption flags. | Create a reusable method that accepts a file path and password, returns load duration, encryption status, and handles errors gracefully. | Explain how to extend the sample to capture memory consumption and CPU usage while opening a password‑protected workbook.

using System;
using System.Diagnostics;
using Aspose.Cells;

// C# sample that loads a password‑protected Excel file using Aspose.Cells LoadOptions, measures the elapsed time with Stopwatch, prints the duration in milliseconds, displays the workbook's encryption flags, and releases resources.
class OpenEncryptedWorkbookPerformance
{
    static void Main()
    {
        // Path to the encrypted workbook and its password
        string filePath = "encrypted.xlsx";
        string password = "password";

        // Prepare load options with the password
        LoadOptions loadOptions = new LoadOptions
        {
            Password = password
        };

        // Measure the time required to open the workbook
        Stopwatch timer = Stopwatch.StartNew();
        Workbook workbook = null;
        try
        {
            workbook = new Workbook(filePath, loadOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to open workbook: {ex.Message}");
            return;
        }
        timer.Stop();

        // Log performance metrics
        Console.WriteLine($"Time to open encrypted workbook: {timer.ElapsedMilliseconds} ms");
        Console.WriteLine($"IsEncrypted (Workbook.Settings): {workbook.Settings.IsEncrypted}");
        Console.WriteLine($"IsWorkbookProtectedWithPassword: {workbook.IsWorkbookProtectedWithPassword}");

        // Clean up
        workbook.Dispose();
    }
}
