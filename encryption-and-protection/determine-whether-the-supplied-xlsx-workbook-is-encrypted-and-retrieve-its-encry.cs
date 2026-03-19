using System;
using Aspose.Cells;

class CheckWorkbookEncryption
{
    static void Main(string[] args)
    {
        // Ensure a file path is provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: CheckWorkbookEncryption <path-to-xlsx>");
            return;
        }

        string filePath = args[0];

        // Detect the file format and encryption status without opening the workbook
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine($"File \"{filePath}\" IsEncrypted (FileFormatInfo): {formatInfo.IsEncrypted}");

        // Demonstrate that the same information is available after loading the workbook
        if (formatInfo.IsEncrypted)
        {
            // If the workbook is encrypted, a password is required to load it.
            // Replace "yourPassword" with the actual password if known.
            string password = "yourPassword";

            LoadOptions loadOptions = new LoadOptions { Password = password };
            try
            {
                Workbook encryptedWb = new Workbook(filePath, loadOptions);
                Console.WriteLine($"Workbook loaded successfully. Settings.IsEncrypted: {encryptedWb.Settings.IsEncrypted}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load encrypted workbook: {ex.Message}");
            }
        }
        else
        {
            // Load normally for an unencrypted workbook
            Workbook wb = new Workbook(filePath);
            Console.WriteLine($"Workbook loaded successfully. Settings.IsEncrypted: {wb.Settings.IsEncrypted}");
        }
    }
}