// Title: Validate that a password‑protected XLSX workbook opens quickly with Aspose.Cells for .NET
// Common Searches: how to measure load time of an encrypted XLSX file using Aspose.Cells in C# | benchmark opening speed of a password protected workbook with Aspose.Cells .NET | test performance of strong password encryption for Excel files in Aspose.Cells | C# code to time opening of a password protected workbook saved as .xlsx | determine if Aspose.Cells can quickly open a workbook encrypted with a strong password
// Tags: Aspose.Cells encrypted workbook load performance | XLSX password encryption opening speed .NET | C# workbook load timing with LoadOptions | strong password Excel encryption validation | measure Aspose.Cells workbook opening latency

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a 1,000‑row by 50‑column XLSX workbook, applies a strong password, saves it, and then reloads it using LoadOptions with the password while timing the operation. It reports the elapsed milliseconds and indicates whether the opening time stays under a 2‑second threshold, helping developers validate that strong encryption does not noticeably degrade load performance in modern Excel.
class WorkbookEncryptionValidation
{
    static void Main()
    {
        // Define a strong password
        const string password = "Str0ngP@ssw0rd!2026";

        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Name = "Data";

        // Populate the worksheet with a reasonable amount of data
        for (int row = 0; row < 1000; row++)
        {
            for (int col = 0; col < 50; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Encrypt the workbook with the strong password
        wb.Settings.Password = password;

        // Save the encrypted workbook to a temporary file
        string encryptedFilePath = "EncryptedWorkbook.xlsx";
        wb.Save(encryptedFilePath, SaveFormat.Xlsx);

        // Prepare load options with the password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = password
        };

        // Measure the time required to open the encrypted workbook
        Stopwatch sw = Stopwatch.StartNew();
        Workbook loadedWb = new Workbook(encryptedFilePath, loadOptions);
        sw.Stop();

        // Output the elapsed time in milliseconds
        Console.WriteLine($"Time to open encrypted workbook: {sw.ElapsedMilliseconds} ms");

        // Simple validation: consider opening fast if under 2000 ms (2 seconds)
        const long fastThresholdMs = 2000;
        if (sw.ElapsedMilliseconds <= fastThresholdMs)
        {
            Console.WriteLine("The encrypted workbook opens quickly in modern Excel versions.");
        }
        else
        {
            Console.WriteLine("Opening the encrypted workbook is slower than expected.");
        }
    }
}
