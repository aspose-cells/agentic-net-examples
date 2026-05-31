using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDetection
{
    class Program
    {
        static void Main()
        {
            // Path to the ODS file to be examined
            string odsFilePath = "sample.ods";

            // Optional: password for the file if it is encrypted.
            // If the file is not encrypted this value is ignored.
            string password = "myPassword";

            // Verify that the file exists before proceeding
            if (!File.Exists(odsFilePath))
            {
                Console.WriteLine($"File not found: {odsFilePath}");
                return;
            }

            try
            {
                // ------------------------------------------------------------
                // 1. Detect basic file format information (including encryption flag)
                // ------------------------------------------------------------
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(odsFilePath);
                Console.WriteLine($"File format detected: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is encrypted (quick check): {formatInfo.IsEncrypted}");

                // ------------------------------------------------------------
                // 2. If the file is encrypted, load it with the supplied password.
                //    Aspose.Cells does not expose the encryption algorithm for ODS files,
                //    so we simply acknowledge that the file is encrypted.
                // ------------------------------------------------------------
                if (formatInfo.IsEncrypted)
                {
                    // LoadOptions allow us to specify the password required to open the file
                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = password
                    };

                    // Load the workbook using the password
                    Workbook workbook = new Workbook(odsFilePath, loadOptions);

                    // Encryption algorithm information is not available for ODS via Aspose.Cells.
                    Console.WriteLine("Encryption algorithm: Not exposed for ODS files.");
                }
                else
                {
                    Console.WriteLine("The ODS file is not encrypted.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}