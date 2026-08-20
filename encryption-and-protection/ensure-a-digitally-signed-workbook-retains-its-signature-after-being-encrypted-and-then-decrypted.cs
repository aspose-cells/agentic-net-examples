// Title: C# – Preserve Digital Signature When Encrypting & Decrypting an Excel Workbook with Aspose.Cells
// Description: Demonstrates how to sign an Excel file using an X509 certificate, apply password protection with Aspose.Cells for .NET, remove the password later, and confirm that the digital signature remains intact and valid.
// Keywords: Aspose.Cells digital signature encryption | C# preserve Excel signature after password protection | remove password from signed workbook Aspose | validate X509 signature after decryption | Excel workbook encryption Aspose.Cells | keep digital signature after encrypting Excel | C# Aspose.Cells LoadOptions password
// Common Searches: how to keep a digital signature after encrypting an Excel file with Aspose.Cells | remove password from a signed workbook without losing the signature | verify Excel digital signature after decryption C# | Aspose.Cells encrypt signed workbook and retain signature | C# example for password‑protecting a digitally signed Excel file
// Developer Intent: Ensure a workbook signed with an X509 certificate stays digitally signed after password‑based encryption and subsequent decryption using Aspose.Cells for .NET.
// Use Cases: Securely transmit a compliance report: sign it, encrypt with a password, then decrypt for audit while the signature stays verifiable. | Archive financially signed statements in encrypted storage and later retrieve them for regulatory review without losing signature integrity. | Automate a workflow that encrypts signed Excel files for backup, then restores them for processing, confirming the original digital signature is still present.
// AI Prompts: Show C# code that encrypts a digitally signed Excel workbook with a password and then removes the password while preserving the signature using Aspose.Cells. | Explain how to check that a digital signature remains valid after decrypting a password‑protected workbook in .NET. | Provide a step‑by‑step guide for handling missing X509 certificates when adding a digital signature but still needing to encrypt and decrypt the file.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureEncryptionDemo
{
    // Demonstrates how to sign an Excel file using an X509 certificate, apply password protection with Aspose.Cells for .NET, remove the password later, and confirm that the digital signature remains intact and valid.
    class Program
    {
        static void Main()
        {
            // Paths (replace with actual paths on your environment)
            string certificatePath = "mycert.pfx";
            string certificatePassword = "certPassword";
            string signedPath = "SignedWorkbook.xlsx";
            string encryptedPath = "EncryptedSignedWorkbook.xlsx";
            string decryptedPath = "DecryptedSignedWorkbook.xlsx";

            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and add some sample data
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Digital Signature Test");

                // -----------------------------------------------------------------
                // 2. Create a digital signature using an X509 certificate (if available)
                // -----------------------------------------------------------------
                if (File.Exists(certificatePath))
                {
                    try
                    {
                        X509Certificate2 cert = new X509Certificate2(certificatePath, certificatePassword);
                        DigitalSignature signature = new DigitalSignature(cert, "Test Signature", DateTime.Now);

                        // -----------------------------------------------------------------
                        // 3. Attach the signature to the workbook
                        // -----------------------------------------------------------------
                        DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                        signatures.Add(signature);
                        workbook.SetDigitalSignature(signatures);
                    }
                    catch (CryptographicException ex)
                    {
                        Console.WriteLine($"Warning: Unable to load certificate. Digital signature will be skipped. Details: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("Warning: Certificate file not found. Digital signature will be skipped.");
                }

                // -----------------------------------------------------------------
                // 4. Save the signed workbook
                // -----------------------------------------------------------------
                workbook.Save(signedPath, SaveFormat.Xlsx);
                if (File.Exists(signedPath))
                {
                    Console.WriteLine($"Signed workbook saved. IsDigitallySigned = {new Workbook(signedPath).IsDigitallySigned}");
                }

                // -----------------------------------------------------------------
                // 5. Encrypt the signed workbook with a password
                // -----------------------------------------------------------------
                workbook.Settings.Password = "encryptionPwd";
                workbook.Save(encryptedPath, SaveFormat.Xlsx);
                if (File.Exists(encryptedPath))
                {
                    Console.WriteLine($"Encrypted workbook saved. IsEncrypted = {new Workbook(encryptedPath, new LoadOptions { Password = "encryptionPwd" }).Settings.IsEncrypted}");
                }

                // -----------------------------------------------------------------
                // 6. Load the encrypted workbook (providing the password)
                // -----------------------------------------------------------------
                LoadOptions loadOptions = new LoadOptions { Password = "encryptionPwd" };
                Workbook encryptedWorkbook = new Workbook(encryptedPath, loadOptions);
                Console.WriteLine($"Loaded encrypted workbook. IsDigitallySigned = {encryptedWorkbook.IsDigitallySigned}");

                // -----------------------------------------------------------------
                // 7. Remove encryption by clearing the password and save again
                // -----------------------------------------------------------------
                encryptedWorkbook.Settings.Password = null; // clears encryption
                encryptedWorkbook.Save(decryptedPath, SaveFormat.Xlsx);
                if (File.Exists(decryptedPath))
                {
                    Console.WriteLine($"Decrypted workbook saved. IsEncrypted = {new Workbook(decryptedPath).Settings.IsEncrypted}");
                }

                // -----------------------------------------------------------------
                // 8. Verify that the digital signature is still present after decryption
                // -----------------------------------------------------------------
                Workbook finalWorkbook = new Workbook(decryptedPath);
                Console.WriteLine($"Final workbook IsDigitallySigned = {finalWorkbook.IsDigitallySigned}");

                // Optional: iterate signatures and display validity
                DigitalSignatureCollection finalSignatures = finalWorkbook.GetDigitalSignature();
                if (finalSignatures != null)
                {
                    foreach (DigitalSignature ds in finalSignatures)
                    {
                        Console.WriteLine($"Signature Comment: {ds.Comments}, IsValid: {ds.IsValid}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
