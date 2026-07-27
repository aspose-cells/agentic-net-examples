// Title: Keep a Digital Signature Intact When Encrypting and Decrypting an Excel Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to sign an Excel workbook using an X509 certificate, protect it with a password, load the encrypted file to verify the signature, remove the password, and confirm that the digital signature remains after decryption using Aspose.Cells for C#.
// Keywords: Aspose.Cells digital signature encryption | C# sign Excel workbook | password protect Excel Aspose.Cells | IsDigitallySigned after decryption | X509Certificate2 Aspose.Cells example | preserve signature encrypted workbook | load encrypted Excel with password | remove password from signed workbook
// Common Searches: how to retain digital signature after password protecting Excel with Aspose.Cells | verify IsDigitallySigned on encrypted workbook .NET | decrypt signed Excel file without losing signature Aspose.Cells | Aspose.Cells example: sign, encrypt, and decrypt workbook | remove password from signed workbook Aspose.Cells C#
// Developer Intent: Ensure a workbook that has been digitally signed stays signed after applying password protection and after the password is removed.
// Use Cases: Create and digitally sign a new workbook, then apply password protection while keeping the signature valid. | Load a password‑protected workbook using LoadOptions, check Workbook.IsDigitallySigned, and confirm the signature is still present. | Remove the password from a signed workbook, save the decrypted file, and verify that the digital signature persists.
// AI Prompts: Generate C# code that signs an Excel file with Aspose.Cells, encrypts it with a password, then decrypts it while preserving the digital signature. | Explain the steps Aspose.Cells performs to maintain a digital signature during password protection and subsequent decryption. | Provide troubleshooting tips if Workbook.IsDigitallySigned returns false after decrypting a previously signed workbook.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureEncryptionDemo
{
    // Demonstrates how to sign an Excel workbook using an X509 certificate, protect it with a password, load the encrypted file to verify the signature, remove the password, and confirm that the digital signature remains after decryption using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Paths (replace with actual paths as needed)
            string certificatePath = "certificate.pfx";
            string certificatePassword = "certpwd";
            string signedPath = "SignedWorkbook.xlsx";
            string encryptedPath = "EncryptedWorkbook.xlsx";
            string decryptedPath = "DecryptedWorkbook.xlsx";

            try
            {
                // -------------------------------------------------
                // 1. Create a workbook and add a digital signature
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Digitally signed content");

                // Verify certificate file exists before loading
                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                    return;
                }

                // Load certificate and create a digital signature
                X509Certificate2 cert = new X509Certificate2(certificatePath, certificatePassword);
                DigitalSignature signature = new DigitalSignature(cert, "Demo Signature", DateTime.Now);
                DigitalSignatureCollection signatures = new DigitalSignatureCollection { signature };

                // Apply the signature to the workbook
                workbook.SetDigitalSignature(signatures);

                // Save the signed workbook
                workbook.Save(signedPath, SaveFormat.Xlsx);
                Console.WriteLine($"Signed workbook saved to: {signedPath}");

                // -------------------------------------------------
                // 2. Verify the workbook is digitally signed
                // -------------------------------------------------
                Workbook signedWorkbook = new Workbook(signedPath);
                Console.WriteLine("Initially signed? " + signedWorkbook.IsDigitallySigned); // Expected: True

                // -------------------------------------------------
                // 3. Encrypt the signed workbook with a password
                // -------------------------------------------------
                signedWorkbook.Settings.Password = "encryptionPwd";
                signedWorkbook.Save(encryptedPath, SaveFormat.Xlsx);
                Console.WriteLine($"Encrypted workbook saved to: {encryptedPath}");

                // -------------------------------------------------
                // 4. Load the encrypted workbook (providing password) and verify signature
                // -------------------------------------------------
                LoadOptions loadOptions = new LoadOptions { Password = "encryptionPwd" };
                Workbook encryptedWorkbook = new Workbook(encryptedPath, loadOptions);
                Console.WriteLine("After encryption, signed? " + encryptedWorkbook.IsDigitallySigned); // Expected: True

                // -------------------------------------------------
                // 5. Decrypt the workbook (remove password) and save
                // -------------------------------------------------
                encryptedWorkbook.Settings.Password = null; // Removing encryption
                encryptedWorkbook.Save(decryptedPath, SaveFormat.Xlsx);
                Console.WriteLine($"Decrypted workbook saved to: {decryptedPath}");

                // -------------------------------------------------
                // 6. Load the decrypted workbook and confirm signature persists
                // -------------------------------------------------
                Workbook decryptedWorkbook = new Workbook(decryptedPath);
                Console.WriteLine("After decryption, signed? " + decryptedWorkbook.IsDigitallySigned); // Expected: True
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
