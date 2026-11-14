// Title: Detect the file format of a decrypted encrypted OOXML (.xlsx) workbook and verify compatibility with Aspose.Cells for .NET
// Common Searches: how to open a password‑protected .xlsx with Aspose.Cells and determine its file format in C# | using Aspose.Cells LoadOptions to decrypt an encrypted Excel file and check the workbook format | C# code to verify that a decrypted OOXML workbook can be saved without errors using Aspose.Cells | detect workbook FileFormat after loading encrypted Excel file with Aspose.Cells .NET | validate compatibility of decrypted Excel file after password decryption in Aspose.Cells
// Tags: Aspose.Cells LoadOptions password decryption | workbook file format detection post‑decryption | verify decrypted OOXML workbook compatibility | C# Aspose.Cells encrypted xlsx handling | FileFormat detection Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an encrypted .xlsx workbook using Aspose.Cells LoadOptions with a password, detects the workbook's FileFormat type, and then saves the decrypted file as a new .xlsx to confirm that decryption succeeded and the resulting OOXML file remains compatible.
class EncryptedOoxmlHandler
{
    static void Main()
    {
        try
        {
            // Path to the encrypted OOXML file (e.g., .xlsx)
            string encryptedFilePath = "encrypted.xlsx";

            // Password used to encrypt the file
            string password = "yourPassword";

            // Verify that the input file exists
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"Input file not found: {encryptedFilePath}");
                return;
            }

            // Load the encrypted workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Detect the file format of the decrypted workbook
            FileFormatType formatType = workbook.FileFormat;
            Console.WriteLine($"Detected format: {formatType}");

            // Save the decrypted workbook (ensures it can be written correctly)
            string decryptedPath = "decrypted.xlsx";
            workbook.Save(decryptedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Decryption and compatibility check completed successfully. Decrypted file saved to '{decryptedPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
