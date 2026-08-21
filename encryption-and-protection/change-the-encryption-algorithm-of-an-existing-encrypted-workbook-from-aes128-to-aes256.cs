// Title: Convert an AES‑128 protected Excel workbook to AES‑256 with Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to open a password‑protected workbook encrypted with AES‑128, re‑apply the same password, switch the encryption to the StrongCryptographicProvider (AES‑256, 256‑bit key), and save the file under a new name. It includes file‑existence checks and exception handling for robust execution.
// Keywords: Aspose.Cells AES-128 to AES-256 | C# re‑encrypt Excel workbook | StrongCryptographicProvider 256 bit | change Excel encryption Aspose | upgrade Excel file security .NET | password protected workbook encryption
// Common Searches: how to change Excel encryption from AES-128 to AES-256 using Aspose.Cells | Aspose.Cells .NET re‑encrypt workbook with stronger algorithm | convert password protected Excel file to AES-256 C# | set StrongCryptographicProvider encryption in Aspose.Cells | upgrade legacy encrypted Excel files to AES-256
// Developer Intent: Re‑encrypt an existing workbook with AES‑256 while keeping the original password.
// Use Cases: Meet PCI‑DSS or GDPR requirements by upgrading legacy AES‑128 reports to AES‑256. | Batch‑process a folder of protected Excel files to strengthen encryption without altering user credentials. | Prepare encrypted workbooks for distribution through a secure corporate portal that mandates AES‑256.
// AI Prompts: Write C# code that loads an AES‑128 encrypted Excel file with Aspose.Cells, changes the encryption to AES‑256, and saves it using the same password. | Create a reusable method (sourcePath, destPath, password) that re‑encrypts any workbook to AES‑256 with Aspose.Cells and returns success status. | Explain how to verify the encryption algorithm after saving and how to handle common errors when converting Excel encryption levels in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to open a password‑protected workbook encrypted with AES‑128, re‑apply the same password, switch the encryption to the StrongCryptographicProvider (AES‑256, 256‑bit key), and save the file under a new name. It includes file‑existence checks and exception handling for robust execution.
class ChangeEncryption
{
    static void Main()
    {
        // Path to the existing AES‑128 encrypted workbook
        string sourcePath = "EncryptedAES128.xlsx";

        // Password used to open the encrypted workbook
        string password = "myPassword";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the workbook with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Preserve the same password after re‑encryption
            workbook.Settings.Password = password;

            // Change encryption to AES‑256 (StrongCryptographicProvider with 256‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the workbook with the new encryption algorithm
            string destPath = "EncryptedAES256.xlsx";
            workbook.Save(destPath);
            Console.WriteLine($"Workbook re‑encrypted and saved to: {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
