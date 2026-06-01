using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionUpgrade
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input parameters:
            // args[0] - source workbook path
            // args[1] - password for the source workbook (empty string if none)
            // args[2] - destination workbook path
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: AsposeCellsEncryptionUpgrade <sourcePath> <password> <destPath>");
                return;
            }

            string sourcePath = args[0];
            string sourcePassword = args[1];
            string destPath = args[2];

            // Detect file format and encryption status without opening the file
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);
            Console.WriteLine($"File format detected: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is encrypted (detected): {formatInfo.IsEncrypted}");

            // Prepare load options – include password only if the file is encrypted
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            if (formatInfo.IsEncrypted)
            {
                loadOptions.Password = sourcePassword;
            }

            // Load the workbook (encrypted or not)
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Verify encryption status after loading
            Console.WriteLine($"Workbook.IsEncrypted after load: {workbook.Settings.IsEncrypted}");

            // If the workbook is encrypted with a legacy algorithm, upgrade it
            if (workbook.Settings.IsEncrypted)
            {
                // Upgrade to the latest encryption standard (StrongCryptographicProvider, 128-bit key)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
                Console.WriteLine("Encryption upgraded to StrongCryptographicProvider (128-bit).");
            }
            else
            {
                Console.WriteLine("Workbook is not encrypted; no upgrade needed.");
            }

            // Save the workbook to the destination path using the same format as the source
            // The Save method automatically chooses the appropriate format based on file extension
            workbook.Save(destPath);
            Console.WriteLine($"Workbook saved to: {destPath}");
        }
    }
}