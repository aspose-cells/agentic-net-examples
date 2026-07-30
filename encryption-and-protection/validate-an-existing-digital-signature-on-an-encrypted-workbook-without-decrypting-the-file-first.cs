// Title: C# – Validate Digital Signatures in a Password‑Protected Excel Workbook with Aspose.Cells
// Description: Shows how to open an encrypted .xlsx using Aspose.Cells LoadOptions, inspect Workbook.IsDigitallySigned, extract the DigitalSignatureCollection, and read each signature’s IsValid flag, comments and signing time—all without manually decrypting the file. Includes handling for missing files, wrong passwords, and absent signatures.
// Keywords: Aspose.Cells | C# digital signature validation | encrypted Excel workbook | password protected XLSX | Workbook.IsDigitallySigned | DigitalSignatureCollection | load encrypted workbook | verify Excel digital signature | Aspose.Cells LoadOptions | signature validation without decryption
// Common Searches: aspnet validate digital signature encrypted excel | how to check if password protected xlsx is signed c# | aspose.cells get digital signatures from encrypted workbook | verify excel digital signature without opening file | c# load encrypted xlsx and read signature info
// Developer Intent: Determine whether a password‑protected Excel file contains valid digital signatures and retrieve their details without manually decrypting the workbook.
// Use Cases: Open an encrypted workbook with a known password and detect if it is digitally signed. | Enumerate all signatures to obtain validation status, comments, and signing timestamps. | Gracefully handle errors such as incorrect passwords, missing files, or workbooks lacking signatures.
// AI Prompts: Generate C# code that uses Aspose.Cells to open a password‑protected .xlsx and return a list of each digital signature’s validity, comments, and sign time. | Explain how Aspose.Cells internally decrypts an encrypted workbook while exposing only signature metadata to the developer. | Suggest modifications to log signature details to a file instead of printing them to the console.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Shows how to open an encrypted .xlsx using Aspose.Cells LoadOptions, inspect Workbook.IsDigitallySigned, extract the DigitalSignatureCollection, and read each signature’s IsValid flag, comments and signing time—all without manually decrypting the file. Includes handling for missing files, wrong passwords, and absent signatures.
class ValidateSignatureEncryptedWorkbook
{
    static void Main()
    {
        // Path to the encrypted workbook that contains a digital signature
        string filePath = "EncryptedSigned.xlsx";

        // Password used to encrypt the workbook
        string password = "yourPassword";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
            return;
        }

        try
        {
            // LoadOptions with the password to open the encrypted file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };

            // Load the workbook (encrypted) using the provided LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Determine whether the workbook is digitally signed
            Console.WriteLine("Workbook is digitally signed: " + workbook.IsDigitallySigned);

            if (workbook.IsDigitallySigned)
            {
                // Retrieve the collection of digital signatures without modifying the workbook
                DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                // Iterate through each signature and display its validation status
                foreach (DigitalSignature signature in signatures)
                {
                    Console.WriteLine("Signature is valid: " + signature.IsValid);
                    Console.WriteLine("Comments: " + signature.Comments);
                    Console.WriteLine("Signed on: " + signature.SignTime);
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or processing
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
