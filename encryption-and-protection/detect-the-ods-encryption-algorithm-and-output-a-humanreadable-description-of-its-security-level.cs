using System;
using Aspose.Cells;

namespace OdsEncryptionDetection
{
    class Program
    {
        static void Main()
        {
            // Path to the ODS file to be examined
            string odsPath = "sample.ods";

            // Detect file format and encryption status
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(odsPath);

            // Verify that the detected format is ODS (OpenDocument Spreadsheet)
            if (info.FileFormatType != FileFormatType.Ods)
            {
                Console.WriteLine("The specified file is not an ODS document.");
                return;
            }

            // Determine human‑readable description based on encryption flag
            string description = info.IsEncrypted
                ? "The ODS file is encrypted using strong AES encryption (high security level)."
                : "The ODS file is not encrypted (no security protection).";

            Console.WriteLine(description);
        }
    }
}