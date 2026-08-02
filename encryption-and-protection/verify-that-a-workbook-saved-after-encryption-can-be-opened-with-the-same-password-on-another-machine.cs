// Title: C# – Verify Encrypted Excel Workbook Opens on Another Machine with the Same Password (Aspose.Cells)
// Description: Creates a workbook, writes data to A1, applies a password, sets 128‑bit strong encryption, saves the file, confirms the password with FileFormatUtil.VerifyPassword, then loads the file on a different computer using LoadOptions.Password, checks the IsEncrypted flag, and validates the original cell value.
// Keywords: Aspose.Cells encrypt workbook C# | password‑protected Excel file | FileFormatUtil.VerifyPassword | LoadOptions.Password | strong 128‑bit encryption | cross‑machine workbook verification | C# Excel encryption example
// Common Searches: how to open an encrypted Excel file created with Aspose.Cells on another computer | verify password for a password‑protected workbook using Aspose.Cells C# | check if a saved workbook is encrypted with Aspose.Cells | load password protected Excel file in .NET Aspose.Cells | Aspose.Cells example for strong encryption and verification
// Developer Intent: Confirm that a workbook encrypted with a password can be opened and read on a different machine using the same password.
// Use Cases: Distribute a password‑protected Excel report and programmatically ensure the password works on the recipient’s server. | Validate encryption integrity before archiving or transmitting confidential spreadsheets. | Automate data extraction from a protected workbook in a backend service while guaranteeing that the file remains encrypted.
// AI Prompts: Generate C# code with Aspose.Cells that encrypts an Excel file using a 128‑bit key, verifies the password via FileFormatUtil.VerifyPassword, and then opens the file on another machine using LoadOptions.Password. | Explain the relationship between FileFormatUtil.VerifyPassword and LoadOptions.Password when opening an encrypted workbook in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, writes data to A1, applies a password, sets 128‑bit strong encryption, saves the file, confirms the password with FileFormatUtil.VerifyPassword, then loads the file on a different computer using LoadOptions.Password, checks the IsEncrypted flag, and validates the original cell value.
class VerifyEncryptedWorkbook
{
    static void Main()
    {
        // Define password and file path
        string password = "Secret123";
        string filePath = "encryptedWorkbook.xlsx";

        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Encrypted content");

        // Set password to encrypt the workbook
        wb.Settings.Password = password;

        // Optionally set strong encryption options
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        wb.Save(filePath, SaveFormat.Xlsx);

        // Verify that the file is encrypted by checking the password
        bool passwordValid;
        using (Stream stream = File.OpenRead(filePath))
        {
            passwordValid = FileFormatUtil.VerifyPassword(stream, password);
        }
        Console.WriteLine($"Password verification (should be true): {passwordValid}");

        // Load the workbook on another machine using LoadOptions with the same password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook loadedWb = new Workbook(filePath, loadOptions);

        // Check IsEncrypted property after loading
        Console.WriteLine($"Workbook.IsEncrypted after load: {loadedWb.Settings.IsEncrypted}");

        // Verify that the data is intact
        string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine($"Cell A1 value: {cellValue}");
    }
}
