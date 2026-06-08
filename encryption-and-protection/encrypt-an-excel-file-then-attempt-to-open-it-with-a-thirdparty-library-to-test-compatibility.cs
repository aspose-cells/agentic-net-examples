using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        try
        {
            // -------------------- Create and encrypt workbook --------------------
            Workbook workbook = new Workbook();                         // create workbook
            Worksheet sheet = workbook.Worksheets[0];                  // get first worksheet
            sheet.Cells["A1"].PutValue("Encrypted Data");              // add sample data
            sheet.Cells["B2"].PutValue(12345);

            // Set password and encryption options (StrongCryptographicProvider, 128‑bit key)
            workbook.Settings.Password = "mySecretPassword";
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedFile = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedFile, SaveFormat.Xlsx);

            // -------------------- Verify encryption with Aspose --------------------
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedFile);
            Console.WriteLine($"Aspose detection - IsEncrypted: {formatInfo.IsEncrypted}");

            // Ensure the file exists before attempting to load it
            if (!File.Exists(encryptedFile))
            {
                Console.WriteLine($"File not found: {encryptedFile}");
                return;
            }

            // Load the encrypted file using Aspose (requires password)
            LoadOptions loadOptions = new LoadOptions { Password = "mySecretPassword" };
            Workbook loadedWorkbook = new Workbook(encryptedFile, loadOptions);
            Console.WriteLine($"Aspose loaded cell A1: {loadedWorkbook.Worksheets[0].Cells["A1"].StringValue}");

            // -------------------- Attempt to open with a third‑party library --------------------
            // The ExcelDataReader library is not referenced in this project.
            // The following block is kept for illustration and will be skipped at runtime.
            try
            {
                Console.WriteLine("Third‑party library test skipped (ExcelDataReader not available).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Third‑party library failed to open encrypted file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}