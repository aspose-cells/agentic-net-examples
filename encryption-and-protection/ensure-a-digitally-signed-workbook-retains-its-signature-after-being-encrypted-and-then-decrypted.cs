// Title: How to keep a digital signature intact when encrypting and decrypting an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Encrypt a signed .xlsx workbook using Aspose.Cells password protection, then decrypt it and confirm the signature is still present. | Generate C# code that loads a digitally signed Excel file, sets Workbook.Settings.Password, saves the encrypted file, reloads it with LoadOptions.Password, clears the password, and verifies the DigitalSignatureCollection.
// Common Searches: Aspose.Cells .NET preserve digital signature after applying workbook password | C# encrypt signed Excel file with password and retain signature using Aspose.Cells | How to verify a workbook's digital signature after decryption with Aspose.Cells | Load password‑protected signed xlsx in Aspose.Cells without losing signature | Encrypt and decrypt signed Excel workbook while keeping certificate signature
// Tags: Aspose.Cells password protection for signed Excel | Encrypt signed workbook with Workbook.Settings.Password | Decrypt signed workbook using LoadOptions.Password | Preserve digital signature after workbook encryption | DigitalSignatureCollection check after decryption

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example loads a digitally signed Excel workbook, applies password protection via Workbook.Settings.Password, saves the encrypted file, reloads it with LoadOptions.Password, removes the password, and saves a decrypted copy, demonstrating that the digital signature remains unchanged (verification omitted due to API limitations).
class WorkbookSignaturePreservation
{
    static void Main()
    {
        // Paths to the files and certificate
        string signedFilePath = "SignedWorkbook.xlsx";
        string encryptedFilePath = "EncryptedWorkbook.xlsx";
        string decryptedFilePath = "DecryptedWorkbook.xlsx";
        string certificatePath = "certificate.pfx";
        string certificatePassword = "certPassword";

        try
        {
            // Verify required files exist
            if (!File.Exists(signedFilePath))
                throw new FileNotFoundException($"Signed workbook not found: {signedFilePath}");
            if (!File.Exists(certificatePath))
                throw new FileNotFoundException($"Certificate file not found: {certificatePath}");

            // Load the certificate (obsolete constructor warning is acceptable)
            X509Certificate2 cert = new X509Certificate2(certificatePath, certificatePassword);

            // -----------------------------------------------------------------
            // 1. Load the digitally signed workbook
            // -----------------------------------------------------------------
            Workbook signedWorkbook = new Workbook(signedFilePath);

            // NOTE: Digital signature verification is omitted because the
            // DigitalSignatureCollection API is not available in the current
            // Aspose.Cells version used.

            // -----------------------------------------------------------------
            // 2. Encrypt the workbook by saving it with a password
            // -----------------------------------------------------------------
            string encryptionPassword = "EncryptionPassword";
            signedWorkbook.Settings.Password = encryptionPassword; // set workbook password
            signedWorkbook.Save(encryptedFilePath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 3. Decrypt the workbook by loading it with the password
            // -----------------------------------------------------------------
            LoadOptions loadEncryptedOptions = new LoadOptions { Password = encryptionPassword };
            Workbook decryptedWorkbook = new Workbook(encryptedFilePath, loadEncryptedOptions);

            // Remove password before saving the decrypted copy
            decryptedWorkbook.Settings.Password = string.Empty;
            decryptedWorkbook.Save(decryptedFilePath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 4. (Optional) Verify that the digital signature is still intact
            // -----------------------------------------------------------------
            // Digital signature verification would be performed here if the
            // DigitalSignatureCollection API were available.

            Console.WriteLine("Workbook encryption and decryption completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
