// Title: Encrypt an Excel workbook with Aspose.Cells .NET and confirm password protection fails without a password
// Description: Creates a new Workbook, writes data to A1, applies a password with StrongCryptographicProvider (128‑bit), saves the file, checks the IsEncrypted flag, then attempts to load the same file without credentials and captures the expected exception that indicates access is denied.
// Keywords: Aspose.Cells encrypt workbook | C# password protect Excel | Workbook.IsEncrypted | StrongCryptographicProvider | Excel encryption example | Aspose.Cells .NET security
// Common Searches: Aspose.Cells encrypt Excel file C# | How to set password on workbook using Aspose.Cells | Verify encrypted workbook cannot be opened without password | Aspose.Cells encryption options example | C# test Excel file password protection
// Developer Intent: Apply a password to an Excel workbook with Aspose.Cells and ensure that opening the file without the password throws an exception.
// Use Cases: Secure confidential spreadsheets before distribution | Automate compliance checks that generated reports are password‑protected | Integrate encryption validation into CI/CD pipelines
// AI Prompts: Write C# code using Aspose.Cells to encrypt a workbook with a custom password and 256‑bit AES encryption, then show how to catch the specific exception when opening it without a password. | Provide a step‑by‑step explanation of retrieving the exception type and message from Aspose.Cells when a protected workbook is accessed without credentials.

using System;
using Aspose.Cells;

// Creates a new Workbook, writes data to A1, applies a password with StrongCryptographicProvider (128‑bit), saves the file, checks the IsEncrypted flag, then attempts to load the same file without credentials and captures the expected exception that indicates access is denied.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Secret Data");

        // Set a password to encrypt the workbook
        wb.Settings.Password = "mySecret";

        // Optionally specify encryption type and key length
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // Confirm that the workbook reports being encrypted
        Console.WriteLine($"Workbook IsEncrypted after save: {wb.Settings.IsEncrypted}");

        // Attempt to open the encrypted workbook without providing a password
        try
        {
            Workbook wbOpen = new Workbook(filePath);
            // If opening succeeds, check the encryption flag (should be true)
            Console.WriteLine($"Opened workbook IsEncrypted: {wbOpen.Settings.IsEncrypted}");
        }
        catch (Exception ex)
        {
            // Expected path: opening fails because password is required
            Console.WriteLine("Failed to open encrypted workbook without password: " + ex.Message);
        }
    }
}
