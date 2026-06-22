using System;
using System.Diagnostics;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path for the encrypted workbook
        string filePath = "encryptedWorkbook.xlsx";

        // -----------------------------------------------------------------
        // Create a new workbook, set a password, and save it (encrypted)
        // -----------------------------------------------------------------
        Workbook workbookToEncrypt = new Workbook();                     // create
        workbookToEncrypt.Settings.Password = "mySecretPassword";       // set encryption password
        workbookToEncrypt.Worksheets[0].Cells["A1"].PutValue("Sample data");
        workbookToEncrypt.Save(filePath);                                // save

        // ---------------------------------------------------------------
        // Measure the time required to open the encrypted workbook
        // ---------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();                    // load options
        loadOptions.Password = "mySecretPassword";

        Stopwatch timer = Stopwatch.StartNew();                         // start timing
        Workbook openedWorkbook = new Workbook(filePath, loadOptions);  // load (open) encrypted file
        timer.Stop();                                                   // stop timing

        // Log performance metrics
        Console.WriteLine($"Time to open encrypted workbook: {timer.ElapsedMilliseconds} ms");
        Console.WriteLine($"Workbook IsEncrypted property: {openedWorkbook.Settings.IsEncrypted}");

        // Clean up
        workbookToEncrypt.Dispose();
        openedWorkbook.Dispose();
    }
}