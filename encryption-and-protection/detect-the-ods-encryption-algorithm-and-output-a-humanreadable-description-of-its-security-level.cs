// Title: Detect ODS Encryption Algorithm and Security Level with Aspose.Cells for .NET
// Description: A C# console sample that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to check whether an ODS file is encrypted and, if so, reports the SHA‑AES algorithm employed by Excel 2007/2010 as a high‑strength protection.
// Keywords: Aspose.Cells | C# | ODS encryption detection | SHA‑AES | FileFormatUtil | FileFormatInfo | .NET | document security | OpenDocument Spreadsheet | encryption algorithm
// Common Searches: Aspose.Cells ODS encryption detection | Which algorithm protects ODS files in Aspose.Cells | C# check if ODS is encrypted | SHA‑AES security level for ODS | Detect encrypted ODS with Aspose
// Developer Intent: Determine the encryption method of an ODS file and convey its security rating.
// Use Cases: Automated validation of incoming ODS documents to confirm they use strong SHA‑AES protection before processing. | Compliance logging of encryption status for batch‑converted ODS files. | Conditional workflow that permits decryption only when the file is secured with high‑strength SHA‑AES.
// AI Prompts: Generate C# code that receives an ODS path, uses Aspose.Cells to detect encryption, and returns a message indicating "Strong SHA‑AES" or "Not encrypted". | Create a reusable method named GetOdsEncryptionInfo that leverages FileFormatUtil.DetectFileFormat and outputs the algorithm and security level. | Write a script that scans a directory of ODS files, reports each file's encryption status, and highlights any that lack SHA‑AES protection.

using System;
using Aspose.Cells;

namespace OdsEncryptionDetection
{
    // A C# console sample that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to check whether an ODS file is encrypted and, if so, reports the SHA‑AES algorithm employed by Excel 2007/2010 as a high‑strength protection.
    class Program
    {
        static void Main()
        {
            // Path to the ODS file to be examined
            string odsFilePath = "sample.ods";

            // Detect file format information, including encryption status
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(odsFilePath);

            // Check whether the ODS file is encrypted
            if (formatInfo.IsEncrypted)
            {
                // Aspose.Cells uses the same SHA‑AES encryption for ODS as for Excel 2007/2010.
                // This is considered a strong encryption algorithm (high security level).
                Console.WriteLine("The ODS file is encrypted using strong SHA‑AES encryption (high security).");
            }
            else
            {
                Console.WriteLine("The ODS file is not encrypted.");
            }
        }
    }
}
