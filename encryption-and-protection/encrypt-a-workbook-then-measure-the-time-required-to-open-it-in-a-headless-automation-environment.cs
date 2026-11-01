// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and benchmark the opening time in a headless environment
// Common Searches: C# how to password‑protect an XLSX file with Aspose.Cells | measure time to open a password‑protected workbook using Aspose.Cells LoadOptions | load encrypted Excel file in a server‑side .NET application without UI | Aspose.Cells example for benchmarking encrypted workbook load latency | use Workbook.Protect to encrypt Excel and test opening performance in headless mode
// Tags: Workbook.Protect password encryption | LoadOptions.Password encrypted file loading | headless workbook load timing .NET | benchmark opening time of protected XLSX | Aspose.Cells encrypted workbook performance testing

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, adds sample data, applies password protection with Workbook.Protect, saves it as an encrypted XLSX, then loads the file in a non‑UI (headless) .NET context using LoadOptions.Password while measuring the elapsed time with a Stopwatch, finally printing the load duration in milliseconds.
class WorkbookEncryptionTiming
{
    static void Main()
    {
        // Path for the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";
        // Password to protect the workbook
        string password = "StrongPassword123";

        try
        {
            // -------------------------
            // Create a new workbook
            // -------------------------
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Encrypt the workbook with a password
            // The Protect method sets the password for opening the file
            workbook.Protect(ProtectionType.All, password);

            // Save the encrypted workbook to disk
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Ensure the file was created before attempting to load it
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found after saving.");

            // ----------------------------------------------
            // Measure time required to open the encrypted file
            // ----------------------------------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Load the workbook in a headless (non-UI) environment using the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password // Provide the password required to open the encrypted file
            };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            sw.Stop();

            // Output the elapsed time in milliseconds
            Console.WriteLine($"Time to open encrypted workbook: {sw.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
