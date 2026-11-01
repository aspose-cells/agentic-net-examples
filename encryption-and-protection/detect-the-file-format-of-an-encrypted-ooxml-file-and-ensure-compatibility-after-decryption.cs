// Title: Detect the format of an encrypted OOXML workbook and save it as an unencrypted XLSX using Aspose.Cells for .NET
// Common Searches: how to open an encrypted .xlsx and get its FileFormat using Aspose.Cells in C# | save a decrypted Excel workbook as an unencrypted XLSX with Aspose.Cells | detect workbook format after loading a password‑protected file in .NET | convert a non‑XLSX workbook to XLSX for compatibility using Aspose.Cells | remove password protection and preserve format when saving Excel with Aspose.Cells
// Tags: LoadOptions password Aspose.Cells | FileFormatType retrieval Aspose.Cells | XLSX conversion for workbook Aspose.Cells | Export decrypted workbook to XLSX Aspose.Cells | Encrypted Excel format identification Aspose.Cells | Compatibility conversion after decryption Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an encrypted OOXML workbook with a password via LoadOptions, determines its FileFormatType, and then saves the workbook as an unencrypted XLSX, converting to XLSX if the original format differs to ensure compatibility.
class Program
{
    static void Main()
    {
        // Path to the encrypted OOXML file and its password
        string encryptedFilePath = "encrypted.xlsx";
        string password = "yourPassword";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(encryptedFilePath))
        {
            Console.WriteLine($"Input file not found: {encryptedFilePath}");
            return;
        }

        try
        {
            // Load the encrypted workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Get the detected file format type of the loaded workbook
            FileFormatType fileFormat = workbook.FileFormat;
            Console.WriteLine($"Detected format: {fileFormat}");

            // Ensure compatibility: if the format is not XLSX, convert to XLSX
            SaveFormat targetFormat = SaveFormat.Xlsx;
            if (fileFormat != FileFormatType.Xlsx)
            {
                // Convert by saving to XLSX format
                workbook.Save("converted.xlsx", targetFormat);
                Console.WriteLine("File converted to XLSX for compatibility.");
            }
            else
            {
                // Save the decrypted workbook without a password to preserve compatibility
                workbook.Save("decrypted.xlsx", targetFormat);
                Console.WriteLine("Decrypted file saved as XLSX.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
