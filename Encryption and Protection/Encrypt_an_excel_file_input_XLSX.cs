using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    public class EncryptExcelFile
    {
        /// <summary>
        /// Encrypts an existing XLSX file with the specified password.
        /// </summary>
        /// <param name="inputPath">Full path to the source XLSX file.</param>
        /// <param name="outputPath">Full path where the encrypted file will be saved.</param>
        /// <param name="password">Password to protect the workbook.</param>
        public static void Encrypt(string inputPath, string outputPath, string password)
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the password for the workbook (this will encrypt the file using Excel's default SHA‑AES algorithm)
            workbook.Settings.Password = password;

            // Optional: specify encryption type and key length for older Excel formats (ignored for .xlsx)
            // workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the workbook; the password set above causes the file to be saved encrypted
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }

        // Example usage
        public static void Main()
        {
            string sourceFile = "input.xlsx";          // Path to the original XLSX file
            string encryptedFile = "encrypted.xlsx";   // Path for the encrypted output
            string password = "MySecretPassword";      // Desired password

            Encrypt(sourceFile, encryptedFile, password);

            // Verify encryption status
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(encryptedFile);
            Console.WriteLine($"Is the file encrypted? {info.IsEncrypted}");
        }
    }
}