// Title: Validate a digital signature in a password‑protected Excel (.xlsx) using Aspose.Cells C#
// Description: Demonstrates how to verify the password of an encrypted workbook with FileFormatUtil.VerifyPassword, load the file via LoadOptions so Aspose.Cells decrypts it internally, detect digital signatures, retrieve the DigitalSignatureCollection, and report each signature's IsValid status.
// Keywords: Aspose.Cells digital signature validation | encrypted Excel workbook C# | verify password without manual decryption | LoadOptions password Excel | DigitalSignatureCollection Aspose | XLSX signature check .NET | password‑protected workbook integrity
// Common Searches: how to check a digital signature in a protected .xlsx with Aspose.Cells | verify Excel file password before loading in .NET | validate signatures in encrypted Excel using C# | Aspose.Cells load encrypted workbook and read signatures | check if encrypted workbook is signed
// Developer Intent: Confirm the authenticity of a digitally signed, password‑protected Excel file without performing explicit decryption steps.
// Use Cases: Ensure compliance by verifying signed encrypted spreadsheets before data extraction. | Automate integrity checks in document‑exchange pipelines that receive protected Excel files. | Integrate signature validation into security‑focused reporting tools.
// AI Prompts: Generate C# code that uses Aspose.Cells to verify a workbook password and then validate all embedded digital signatures. | Explain the internal decryption process of Aspose.Cells when a password is supplied via LoadOptions and how to access the DigitalSignatureCollection. | Provide best‑practice error handling for signature validation in password‑protected Excel files with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureValidation
{
    // Demonstrates how to verify the password of an encrypted workbook with FileFormatUtil.VerifyPassword, load the file via LoadOptions so Aspose.Cells decrypts it internally, detect digital signatures, retrieve the DigitalSignatureCollection, and report each signature's IsValid status.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted workbook that contains a digital signature
            string filePath = "EncryptedSignedWorkbook.xlsx";

            // Password used to protect the workbook
            string password = "yourPassword";

            try
            {
                // Verify that the file exists before attempting any operations
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Verify that the supplied password is correct without decrypting the file manually
                bool isPasswordValid;
                using (FileStream stream = File.OpenRead(filePath))
                {
                    isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
                }
                Console.WriteLine($"Password valid: {isPasswordValid}");

                if (!isPasswordValid)
                {
                    Console.WriteLine("Incorrect password. Cannot validate signature.");
                    return;
                }

                // Load the workbook using the password (Aspose.Cells handles decryption internally)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Check if the workbook is digitally signed
                if (workbook.IsDigitallySigned)
                {
                    // Retrieve the collection of digital signatures
                    DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                    // Iterate through each signature and output its validation status
                    foreach (DigitalSignature signature in signatures)
                    {
                        Console.WriteLine($"Signature valid: {signature.IsValid}");
                    }
                }
                else
                {
                    Console.WriteLine("Workbook is not digitally signed.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
