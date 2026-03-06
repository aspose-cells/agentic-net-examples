using System;
using Aspose.Cells;
using Aspose.Cells.Ods;
using Aspose.Cells.Utility;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for source, encrypted and decrypted files
            string sourceXlsx = "input.xlsx";               // original XLSX file
            string encryptedOds = "encrypted_output.ods";   // ODS file saved with password
            string decryptedXlsx = "decrypted_output.xlsx"; // XLSX file after decryption

            // Password to protect the workbook
            const string password = "MySecretPassword";

            // -------------------------------------------------
            // 1. Load the source XLSX workbook (no password needed)
            // -------------------------------------------------
            Workbook workbook = new Workbook(sourceXlsx);

            // -------------------------------------------------
            // 2. Apply password protection and encryption options
            // -------------------------------------------------
            workbook.Settings.Password = password; // set the password
            // Optional: specify encryption algorithm and key length
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // -------------------------------------------------
            // 3. Save the workbook as ODS with encryption
            // -------------------------------------------------
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.GeneratorType = OdsGeneratorType.LibreOffice; // any generator type is fine
            workbook.Save(encryptedOds, saveOptions);

            Console.WriteLine($"Encrypted ODS file saved to: {encryptedOds}");

            // -------------------------------------------------
            // 4. Verify that the saved ODS file is encrypted
            // -------------------------------------------------
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedOds);
            Console.WriteLine($"Is the ODS file encrypted? {formatInfo.IsEncrypted}");

            // -------------------------------------------------
            // 5. Load the encrypted ODS file using the password
            // -------------------------------------------------
            OdsLoadOptions loadOptions = new OdsLoadOptions();
            loadOptions.Password = password; // provide password for decryption

            Workbook encryptedWorkbook = new Workbook(encryptedOds, loadOptions);
            Console.WriteLine("Encrypted ODS file loaded successfully.");

            // -------------------------------------------------
            // 6. Remove password protection (optional) and save as XLSX
            // -------------------------------------------------
            encryptedWorkbook.Settings.Password = null; // clear password
            encryptedWorkbook.Save(decryptedXlsx, SaveFormat.Xlsx);
            Console.WriteLine($"Decrypted workbook saved to: {decryptedXlsx}");
        }
    }
}