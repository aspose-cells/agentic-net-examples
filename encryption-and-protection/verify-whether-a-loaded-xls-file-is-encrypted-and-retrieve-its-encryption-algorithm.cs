using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    class Program
    {
        static void Main()
        {
            // Path to the XLS file to be examined
            string filePath = "sample.xls";

            // Detect file format and encryption status without loading the workbook
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
            Console.WriteLine($"Is the file encrypted? {fileInfo.IsEncrypted}");

            // If the file is encrypted, attempt to load it with a password
            if (fileInfo.IsEncrypted)
            {
                // Prepare load options with the password (replace with actual password)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = "yourPassword"
                };

                // Load the workbook using the provided password
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Verify encryption status via workbook settings
                Console.WriteLine($"Workbook.Settings.IsEncrypted: {workbook.Settings.IsEncrypted}");

                // NOTE: Aspose.Cells does not expose the specific encryption algorithm directly.
                // The encryption algorithm used by Excel (e.g., XOR, AES) is handled internally.
                // If needed, you can infer the algorithm based on the file format version
                // (XLS uses older algorithms, XLSX uses AES). Here we simply report the status.
            }
        }
    }
}