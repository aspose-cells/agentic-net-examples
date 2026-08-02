// Title: Validate Digital Signatures in an Encrypted Excel Workbook with Aspose.Cells (C#)
// Description: C# example that verifies the password of a protected .xlsx using FileFormatUtil.VerifyPassword, loads the workbook with LoadOptions.Password, checks Workbook.IsDigitallySigned, retrieves the DigitalSignatureCollection, and reports each signature's IsValid status.
// Keywords: Aspose.Cells | C# | digital signature validation | encrypted workbook | password‑protected Excel | FileFormatUtil.VerifyPassword | LoadOptions.Password | Workbook.IsDigitallySigned | GetDigitalSignature | DigitalSignature.IsValid | Excel encryption example
// Common Searches: validate digital signature in password protected Excel using Aspose.Cells | verify workbook password without decrypting Aspose.Cells C# | check if encrypted .xlsx is digitally signed | retrieve and validate signatures from encrypted Excel file | Aspose.Cells example for encrypted workbook signature verification
// Developer Intent: Confirm that an existing digital signature remains valid in an encrypted Excel file without manually decrypting the workbook.
// Use Cases: Quickly test whether the supplied password unlocks the workbook before processing. | Determine if a protected workbook contains any digital signatures. | Iterate through all signatures and evaluate their IsValid property to ensure integrity.
// AI Prompts: Generate C# code that uses Aspose.Cells to verify the password of an encrypted .xlsx and then validate its digital signatures. | Explain step‑by‑step how to load a password‑protected workbook, detect signatures, and check each signature's validity with Aspose.Cells. | Provide best‑practice error handling for signature validation in encrypted Excel files using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace Example
{
    // C# example that verifies the password of a protected .xlsx using FileFormatUtil.VerifyPassword, loads the workbook with LoadOptions.Password, checks Workbook.IsDigitallySigned, retrieves the DigitalSignatureCollection, and reports each signature's IsValid status.
    class ValidateSignatureEncryptedWorkbook
    {
        static void Main()
        {
            // Path to the encrypted workbook and its password
            string filePath = "EncryptedSigned.xlsx";
            string password = "yourPassword";

            // Ensure the file exists before attempting to open it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Verify the password without fully decrypting the file
                using (FileStream stream = File.OpenRead(filePath))
                {
                    bool isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, password);
                    Console.WriteLine($"Password correct: {isPasswordCorrect}");
                }

                // Load the workbook using the password (Aspose.Cells handles decryption internally)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Check if the workbook contains a digital signature
                Console.WriteLine($"Workbook is digitally signed: {workbook.IsDigitallySigned}");

                if (workbook.IsDigitallySigned)
                {
                    // Retrieve the digital signatures collection
                    DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                    // Iterate through each signature and display its validation status
                    foreach (DigitalSignature signature in signatures)
                    {
                        Console.WriteLine($"Signature valid: {signature.IsValid}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
