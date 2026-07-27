using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the ODS file to be examined
            string odsPath = "sample.ods";

            // Detect file format information, including encryption flag
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(odsPath);

            // Determine if the file is encrypted
            bool isEncrypted = formatInfo.IsEncrypted;

            // ODS files use standard ODF encryption (AES). If not encrypted, report none.
            string algorithmName = isEncrypted ? "AES (standard ODS encryption)" : "None";

            Console.WriteLine($"Is the ODS file encrypted? {isEncrypted}");
            Console.WriteLine($"Encryption algorithm: {algorithmName}");

            // Optional: load the workbook to verify the Settings.IsEncrypted property
            // (requires correct password if the file is encrypted)
            LoadOptions loadOptions = new LoadOptions();
            if (isEncrypted)
            {
                // Replace with the actual password if known
                loadOptions.Password = "your_password";
            }

            Workbook workbook = new Workbook(odsPath, loadOptions);
            Console.WriteLine($"Workbook.Settings.IsEncrypted: {workbook.Settings.IsEncrypted}");
        }
    }
}