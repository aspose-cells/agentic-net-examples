using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create and encrypt ODS file --------------------
            // Create a new workbook (default format is XLSX, but we will save as ODS)
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Set a password to protect the workbook (this will encrypt the file)
            workbook.Settings.Password = "MySecretPassword";

            // Save the workbook as ODS (OpenDocument Spreadsheet)
            string encryptedOdsPath = "EncryptedDocument.ods";
            workbook.Save(encryptedOdsPath, SaveFormat.ODS);

            // Verify that the file is indeed encrypted
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedOdsPath);
            Console.WriteLine($"Is the ODS file encrypted? {formatInfo.IsEncrypted}");

            // -------------------- Load and decrypt ODS file --------------------
            // Prepare load options with the password used during encryption
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "MySecretPassword";

            // Load the encrypted ODS file using the password
            Workbook loadedWorkbook = new Workbook(encryptedOdsPath, loadOptions);

            // Access the data to confirm successful decryption
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Decrypted cell value: {cellValue}");

            // -------------------- Optional: Remove protection (decrypt) --------------------
            // Clear the password to save an unprotected version
            loadedWorkbook.Settings.Password = null;
            string decryptedOdsPath = "DecryptedDocument.ods";
            loadedWorkbook.Save(decryptedOdsPath, SaveFormat.ODS);
            Console.WriteLine($"Decrypted ODS file saved to: {decryptedOdsPath}");
        }
    }
}