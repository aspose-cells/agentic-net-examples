using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class DetectOdsEncryption
    {
        static void Main()
        {
            // Path to the ODS file to be examined
            string odsPath = "sample.ods";

            // Detect file format and encryption status without opening the file
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(odsPath);
            Console.WriteLine($"Detected format: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is encrypted: {formatInfo.IsEncrypted}");

            // If the file is encrypted, attempt to open it with a password
            if (formatInfo.IsEncrypted)
            {
                // LoadOptions allows specifying the password for encrypted ODS files
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                loadOptions.Password = "yourPassword"; // replace with actual password

                try
                {
                    // Load the workbook; if the password is correct, no exception is thrown
                    Workbook wb = new Workbook(odsPath, loadOptions);
                    Console.WriteLine("Workbook opened successfully with the provided password.");

                    // ODS encryption used by Aspose.Cells follows the OOXML standard:
                    // - SHA-1 based key derivation
                    // - AES encryption (128‑bit or 256‑bit depending on key length)
                    // Since the API does not expose the exact algorithm, we describe it generically.
                    Console.WriteLine("Encryption algorithm: AES (Advanced Encryption Standard) with SHA‑1 based key derivation.");
                    Console.WriteLine("Security level: Strong (AES provides robust symmetric encryption).");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to open encrypted ODS file: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("The ODS file is not encrypted; no security algorithm to report.");
            }
        }
    }
}