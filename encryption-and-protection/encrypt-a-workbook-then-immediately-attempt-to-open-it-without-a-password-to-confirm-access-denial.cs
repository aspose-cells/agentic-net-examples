// Title: Encrypt an Excel workbook with Aspose.Cells for .NET and confirm password protection
// Description: Creates a workbook, writes confidential data, applies a password with optional AES encryption, saves the file, checks the IsEncrypted flag, attempts to open it without a password (expecting failure), and finally opens it with LoadOptions using the correct password.
// Keywords: Aspose.Cells encrypt workbook C# | password protect Excel .NET | SetEncryptionOptions Aspose.Cells | Workbook.IsEncrypted property | LoadOptions password protected Excel | verify encrypted workbook access
// Common Searches: How to password protect an Excel file using Aspose.Cells in C# | Check if an Excel workbook is encrypted with Aspose.Cells | Open encrypted Excel file with LoadOptions Aspose.Cells | Fail to load password‑protected workbook without password
// Developer Intent: Apply a password to an Excel workbook, save it, and ensure it cannot be opened without the correct password.
// Use Cases: Secure sensitive spreadsheets before sharing with clients or partners. | Automate compliance checks that all generated workbooks are encrypted. | Add a validation step in CI/CD pipelines to reject unprotected Excel outputs.
// AI Prompts: Write C# code using Aspose.Cells to encrypt a workbook with AES‑256 and handle the exception when opened without a password. | Explain how to use LoadOptions to open a password‑protected Excel file and read a specific cell value. | Suggest a logging approach for recording encryption status and unauthorized access attempts when loading encrypted workbooks.

using System;
using Aspose.Cells;

// Creates a workbook, writes confidential data, applies a password with optional AES encryption, saves the file, checks the IsEncrypted flag, attempts to open it without a password (expecting failure), and finally opens it with LoadOptions using the correct password.
class Program
{
    static void Main()
    {
        // Create a new workbook and put some data in it
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive Information");

        // Encrypt the workbook with a password
        workbook.Settings.Password = "StrongPassword123";

        // Optional: define encryption algorithm and key length
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook to disk
        string encryptedFile = "EncryptedWorkbook.xlsx";
        workbook.Save(encryptedFile, SaveFormat.Xlsx);

        // Verify that the workbook reports as encrypted
        Console.WriteLine($"Workbook.IsEncrypted after save: {workbook.Settings.IsEncrypted}");

        // Attempt to open the encrypted workbook without providing a password
        try
        {
            Workbook withoutPassword = new Workbook(encryptedFile);
            // If loading succeeds, check the encryption flag (should be true)
            Console.WriteLine($"Loaded workbook IsEncrypted: {withoutPassword.Settings.IsEncrypted}");
        }
        catch (Exception ex)
        {
            // Expected path: loading fails because password is required
            Console.WriteLine($"Failed to open encrypted workbook without password: {ex.Message}");
        }

        // Load the workbook correctly using the password to demonstrate successful access
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "StrongPassword123";
        Workbook withPassword = new Workbook(encryptedFile, loadOptions);
        Console.WriteLine($"Successfully opened with password. Cell A1 value: {withPassword.Worksheets[0].Cells["A1"].StringValue}");
    }
}
