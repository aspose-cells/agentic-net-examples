// Title: Check if an Excel workbook in a MemoryStream is encrypted with Aspose.Cells for .NET
// Description: Demonstrates using Aspose.Cells' FileFormatUtil to inspect a MemoryStream for encryption, loading the workbook with a password when required, and confirming the flag via Workbook.Settings.IsEncrypted.
// Keywords: Aspose.Cells | C# | MemoryStream | encrypted workbook detection | FileFormatUtil | DetectFileFormat | Workbook.Settings.IsEncrypted | password‑protected Excel | load encrypted Excel | .NET Excel encryption check
// Common Searches: Aspose.Cells detect password protected Excel from stream | C# determine if Excel file is encrypted before loading | How to check encryption of workbook in MemoryStream | Load encrypted Excel using Aspose.Cells LoadOptions | FileFormatUtil.IsEncrypted example
// Developer Intent: Identify the encryption state of a workbook supplied as a stream and open it with the appropriate password only when required.
// Use Cases: Screen user‑uploaded Excel files for password protection before processing | Automatically apply stored passwords when reading encrypted workbooks from network streams | Log encryption status of in‑memory Excel files for compliance audits | Skip decryption steps for unprotected files to improve performance | Provide a fallback for opening both encrypted and plain workbooks in a single routine
// AI Prompts: Write a C# method that receives a MemoryStream and returns true if the stream contains an encrypted Excel workbook using Aspose.Cells. | Explain why FileFormatUtil.DetectFileFormat.IsEncrypted may differ from Workbook.Settings.IsEncrypted after loading. | Generate sample code that reads an encrypted workbook from a stream with a password supplied from configuration. | Create a try‑catch pattern for handling incorrect passwords when opening encrypted Excel files with Aspose.Cells. | Provide a PowerShell script that calls a .NET assembly to check Excel encryption status.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    // Demonstrates using Aspose.Cells' FileFormatUtil to inspect a MemoryStream for encryption, loading the workbook with a password when required, and confirming the flag via Workbook.Settings.IsEncrypted.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file (could be encrypted or not)
            string filePath = "sample.xlsx";

            // Load the file into a memory stream
            using (MemoryStream memoryStream = new MemoryStream(File.ReadAllBytes(filePath)))
            {
                // Detect file format information directly from the stream
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(memoryStream);
                Console.WriteLine($"Is the workbook encrypted (detected from stream)? {formatInfo.IsEncrypted}");

                // Reset stream position before any further operations
                memoryStream.Position = 0;

                // If the workbook is encrypted, load it with a password (replace with actual password)
                Workbook workbook;
                if (formatInfo.IsEncrypted)
                {
                    // Example password; replace with the correct one for your file
                    string password = "yourPassword";

                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = password
                    };
                    workbook = new Workbook(memoryStream, loadOptions);
                }
                else
                {
                    // Load normally when not encrypted
                    workbook = new Workbook(memoryStream);
                }

                // After loading, you can also check the Settings.IsEncrypted property
                Console.WriteLine($"Is the loaded workbook encrypted (Workbook.Settings.IsEncrypted)? {workbook.Settings.IsEncrypted}");
            }
        }
    }
}
