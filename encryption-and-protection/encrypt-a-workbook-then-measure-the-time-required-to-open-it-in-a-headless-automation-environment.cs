// Title: Measure the time required to open a password‑protected XLSX workbook with Aspose.Cells in a headless C# automation test
// AI Prompts: Create a C# console program that generates an XLSX file, sets Workbook.Settings.Password, saves it, then reloads the file using LoadOptions.Password while measuring the load duration with Stopwatch. | Modify the program to read the file path and password from command‑line arguments and print the opening time in milliseconds. | Add a loop that opens the encrypted workbook repeatedly, records each elapsed time, and outputs the average load time for performance testing in a headless environment.
// Common Searches: C# Aspose.Cells how to benchmark opening speed of an encrypted Excel workbook | measure load time of password protected XLSX using LoadOptions.Password in .NET | headless automation test for decrypting Excel files with Aspose.Cells | performance testing of workbook decryption latency Aspose.Cells C#
// Tags: measure encrypted workbook load time Aspose.Cells | Aspose.Cells password protection performance benchmark | C# load encrypted XLSX with LoadOptions | headless automation timing workbook decryption | benchmark workbook decryption latency .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Demonstrates creating an Excel workbook, applying password protection via Workbook.Settings.Password, saving it as XLSX, then loading it with LoadOptions.Password while timing the operation using Stopwatch to evaluate decryption performance in a headless C# automation scenario.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Define encryption password
            const string password = "SecretPassword";

            // Apply password protection to the workbook
            workbook.Settings.Password = password;

            // Save the encrypted workbook to disk
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // Verify the file exists before attempting to load it
            if (!File.Exists(encryptedPath))
                throw new FileNotFoundException("Encrypted workbook file not found.", encryptedPath);

            // Prepare load options with the same password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            // Measure the time required to open the encrypted workbook
            Stopwatch stopwatch = Stopwatch.StartNew();
            Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);
            stopwatch.Stop();

            // Output the elapsed time
            Console.WriteLine($"Time to open encrypted workbook: {stopwatch.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
