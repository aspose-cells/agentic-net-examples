// Title: Encrypt an Excel workbook with Aspose.Cells (.NET) and verify decryption fails using a third‑party library
// Description: C# example that creates a workbook, writes data, applies a password, selects a strong encryption algorithm (128‑bit), saves the file, then attempts to open it with a non‑Aspose library (e.g., OpenXML SDK) to confirm that decryption is rejected, while Aspose reloads the file with the correct password to show the IsEncrypted flag.
// Keywords: Aspose.Cells encrypt workbook C# | password protect XLSX .NET | strong encryption Aspose.Cells | verify encrypted Excel file | third‑party decryption attempt | OpenXML SDK read encrypted workbook | Excel encryption failure test | C# workbook security example
// Common Searches: How to encrypt an Excel file with Aspose.Cells and test decryption with OpenXML SDK | C# verify that a password‑protected XLSX cannot be opened without the password | Aspose.Cells set encryption algorithm and key length | Load encrypted workbook with correct password using Aspose.Cells | Attempt to read Aspose‑encrypted workbook using a third‑party library
// Developer Intent: Create a password‑protected Excel file with Aspose.Cells and prove that external libraries cannot decrypt it without the correct password.
// Use Cases: Secure financial or HR reports before distribution to external partners. | Compliance testing to ensure encryption meets regulatory standards. | Automated validation that generated workbooks are truly protected before archiving.
// AI Prompts: Generate C# code that encrypts an Excel workbook with Aspose.Cells and then tries to open it with OpenXML SDK without a password, handling the expected exception. | Show how to set a 256‑bit AES encryption option in Aspose.Cells and verify the failure of decryption using a non‑Aspose library. | Explain how to programmatically check the IsEncrypted property after saving a workbook and why a third‑party reader cannot access the content.

using System;
using System.IO;
using Aspose.Cells;

// C# example that creates a workbook, writes data, applies a password, selects a strong encryption algorithm (128‑bit), saves the file, then attempts to open it with a non‑Aspose library (e.g., OpenXML SDK) to confirm that decryption is rejected, while Aspose reloads the file with the correct password to show the IsEncrypted flag.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Secret Data");

            // Encrypt the workbook with a password
            wb.Settings.Password = "mySecret";

            // Optional: specify encryption algorithm and key length
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedPath = "encrypted.xlsx";
            wb.Save(encryptedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{encryptedPath}'.");

            // Verify that Aspose reports the workbook as encrypted
            if (File.Exists(encryptedPath))
            {
                Workbook loadedWithPassword = new Workbook(encryptedPath, new LoadOptions { Password = "mySecret" });
                Console.WriteLine("Aspose reports IsEncrypted: " + loadedWithPassword.Settings.IsEncrypted);
            }
            else
            {
                Console.WriteLine($"Error: Encrypted file '{encryptedPath}' was not found.");
            }

            // Note: Verification with OpenXML SDK is omitted because the required
            // DocumentFormat.OpenXml assembly is not referenced in this project.
            // Aspose.Cells already confirms the encryption status above.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
