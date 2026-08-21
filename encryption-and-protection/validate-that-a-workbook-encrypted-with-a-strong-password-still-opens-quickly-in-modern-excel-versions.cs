// Title: Benchmark opening time of an AES‑256 encrypted Excel workbook with Aspose.Cells for .NET
// Description: This example creates a workbook, writes sample data, applies a strong password with AES‑256 encryption (StrongCryptographicProvider), saves the file, confirms the encryption flag, loads the workbook using the password while measuring the elapsed time, and validates the encryption status and password via FileFormatUtil.
// Keywords: Aspose.Cells AES 256 encryption | C# benchmark encrypted Excel load time | password protected workbook performance | FileFormatUtil verify password | detect encrypted Excel file | StrongCryptographicProvider | Excel encryption speed .NET | measure workbook opening time
// Common Searches: how to encrypt Excel file with AES‑256 using Aspose.Cells | measure load time of password‑protected workbook in C# | check if Excel file is encrypted with Aspose.Cells | benchmark opening speed of encrypted .xlsx | verify password of encrypted Excel using Aspose.Cells
// Developer Intent: Confirm that an Excel workbook encrypted with a strong password opens quickly when loaded with Aspose.Cells.
// Use Cases: Create and save an Excel workbook with AES‑256 encryption and a strong password. | Load the encrypted workbook with the correct password while timing the operation. | Validate encryption status and password correctness using FileFormatUtil.
// AI Prompts: Generate C# code that logs the opening duration of an AES‑256 encrypted workbook and compares it to a configurable performance threshold. | Show how to encrypt the same workbook with different EncryptionType values and produce a table of opening times for each. | Provide a method that retries opening an encrypted workbook with exponential back‑off when the initial load exceeds a specified time limit.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// This example creates a workbook, writes sample data, applies a strong password with AES‑256 encryption (StrongCryptographicProvider), saves the file, confirms the encryption flag, loads the workbook using the password while measuring the elapsed time, and validates the encryption status and password via FileFormatUtil.
class ValidateEncryptedWorkbook
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Sensitive data");

        // Apply a strong password to the workbook
        wb.Settings.Password = "Str0ngP@ssw0rd!";

        // Set strong encryption options (AES 256-bit)
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";
        wb.Save(filePath);

        // Confirm that the workbook reports being encrypted
        Console.WriteLine("Workbook Settings.IsEncrypted: " + wb.Settings.IsEncrypted);

        // Load the encrypted workbook with the password and measure opening time
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "Str0ngP@ssw0rd!";

        Stopwatch sw = Stopwatch.StartNew();
        Workbook loadedWb = new Workbook(filePath, loadOptions);
        sw.Stop();

        Console.WriteLine("Time to open encrypted workbook (ms): " + sw.ElapsedMilliseconds);

        // Verify encryption status via FileFormatInfo
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine("FileFormatInfo.IsEncrypted: " + formatInfo.IsEncrypted);

        // Validate the password using FileFormatUtil.VerifyPassword
        using (Stream stream = File.OpenRead(filePath))
        {
            bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, "Str0ngP@ssw0rd!");
            Console.WriteLine("Password verification (FileFormatUtil): " + isPasswordValid);
        }
    }
}
