using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionUpgrade
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input parameters
            // args[0] - path to the source workbook
            // args[1] - password to open the workbook (empty string if not password protected)
            // args[2] - path for the upgraded workbook
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: AsposeCellsEncryptionUpgrade <sourcePath> <password> <outputPath>");
                return;
            }

            string sourcePath = args[0];
            string password = args[1];
            string outputPath = args[2];

            // Detect file format and encryption status
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);
            Console.WriteLine($"File format detected: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is encrypted (legacy detection): {formatInfo.IsEncrypted}");

            // If the workbook is not encrypted, simply copy it (or inform the user)
            if (!formatInfo.IsEncrypted)
            {
                Console.WriteLine("Workbook is not encrypted. No upgrade required.");
                // Optionally copy the file as‑is
                System.IO.File.Copy(sourcePath, outputPath, true);
                Console.WriteLine($"File copied to {outputPath}");
                return;
            }

            // Load the encrypted workbook using the supplied password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.Password = password;
            Workbook workbook = new Workbook(sourcePath, loadOptions);
            Console.WriteLine("Encrypted workbook loaded successfully.");

            // Upgrade encryption to the latest standard (StrongCryptographicProvider, 256‑bit key)
            // First, set the password again (required before applying encryption options)
            workbook.Settings.Password = password;
            // Apply new encryption options
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
            Console.WriteLine("Encryption upgraded to StrongCryptographicProvider with 256‑bit key.");

            // Save the upgraded workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Upgraded workbook saved to {outputPath}");

            // Verify upgrade by re‑detecting encryption status
            FileFormatInfo upgradedInfo = FileFormatUtil.DetectFileFormat(outputPath);
            Console.WriteLine($"Upgraded file IsEncrypted: {upgradedInfo.IsEncrypted}");
        }
    }
}