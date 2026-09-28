// Title: Detect and preserve the original format of an encrypted OOXML workbook when decrypting with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells LoadOptions with a password to open an encrypted .xlsx/.xlsm/.xlsb file, read its Workbook.FileFormat property, and save the decrypted workbook using the matching SaveFormat. | Write C# code that verifies an encrypted Excel file exists, decrypts it via Aspose.Cells, determines whether it is Xlsx, Xlsm, or Xlsb, and writes the decrypted copy in the same format.
// Common Searches: how to determine original Excel format of a password protected file using Aspose.Cells | save decrypted workbook in the same format as the encrypted file Aspose.Cells .NET | detect FileFormatType of an encrypted .xlsm file with Aspose.Cells | Aspose.Cells load encrypted OOXML workbook and preserve file type on save | C# decrypt password protected Excel and keep original format
// Tags: detect workbook file format Aspose.Cells | preserve original Excel format after decryption | load password‑protected OOXML workbook C# | map FileFormatType to SaveFormat Aspose.Cells | decrypt encrypted Excel file Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the encrypted Excel file exists, loads it with a password using Aspose.Cells LoadOptions, reads the Workbook.FileFormat property to identify the original OOXML format (Xlsx, Xlsm, or Xlsb), maps this to the appropriate SaveFormat, and saves the decrypted workbook while preserving the original file type.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the encrypted OOXML file
            string encryptedFilePath = "encrypted.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"Error: The file \"{encryptedFilePath}\" was not found.");
                return;
            }

            // Password used to encrypt the file
            string password = "yourPassword";

            // Load the encrypted workbook using LoadOptions with the password.
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            // Load the workbook (decryption happens internally).
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Detect the file format of the loaded workbook.
            FileFormatType detectedFormat = workbook.FileFormat;
            Console.WriteLine($"Detected workbook format: {detectedFormat}");

            // Choose a matching SaveFormat based on the detected FileFormatType.
            SaveFormat saveFormat;
            switch (detectedFormat)
            {
                case FileFormatType.Xlsx:
                    saveFormat = SaveFormat.Xlsx;
                    break;
                case FileFormatType.Xlsm:
                    saveFormat = SaveFormat.Xlsm;
                    break;
                case FileFormatType.Xlsb:
                    saveFormat = SaveFormat.Xlsb;
                    break;
                default:
                    // Fallback to XLSX if the format is unexpected.
                    saveFormat = SaveFormat.Xlsx;
                    break;
            }

            // Save the decrypted workbook using the same format.
            string decryptedFilePath = "decrypted.xlsx";

            try
            {
                workbook.Save(decryptedFilePath, saveFormat);
                Console.WriteLine($"Decrypted workbook saved to: {decryptedFilePath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
