// Title: Detect Excel Workbook Encryption from a MemoryStream with Aspose.Cells for .NET
// Description: Read an Excel file into a MemoryStream, call Aspose.Cells.FileFormatUtil.DetectFileFormat to obtain the IsEncrypted flag, and optionally load the workbook to confirm Workbook.Settings.IsEncrypted.
// Keywords: Aspose.Cells | C# encryption detection | FileFormatUtil | DetectFileFormat | IsEncrypted | MemoryStream | Excel password protection | Workbook.Settings.IsEncrypted | encrypted workbook check | load Excel from stream
// Common Searches: Aspose.Cells detect encrypted Excel file | Check if Excel workbook is password protected in C# | FileFormatUtil DetectFileFormat encryption status | Read Excel from MemoryStream and verify encryption | Determine workbook encryption before opening
// Developer Intent: Identify whether a workbook loaded from a MemoryStream is encrypted without fully opening it.
// Use Cases: Validate encryption of user‑uploaded Excel files to decide if a password prompt is required. | Skip decryption steps for non‑encrypted workbooks, improving processing speed. | Log encryption status of Excel streams received from APIs for compliance reporting.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect encryption of an Excel file from a MemoryStream and handle both encrypted and clear cases. | Show how to catch the exception thrown when loading an encrypted workbook and retrieve the needed password with Aspose.Cells. | Explain the workflow: DetectFileFormat → IsEncrypted flag → conditional Workbook loading → Workbook.Settings.IsEncrypted verification.

using System;
using System.IO;
using Aspose.Cells;

// Read an Excel file into a MemoryStream, call Aspose.Cells.FileFormatUtil.DetectFileFormat to obtain the IsEncrypted flag, and optionally load the workbook to confirm Workbook.Settings.IsEncrypted.
class DetectEncryptionFromMemoryStream
{
    static void Main()
    {
        // Path to the Excel file (encrypted or not)
        string filePath = "sample.xlsx";

        // Load the file into a memory stream
        using (MemoryStream memoryStream = new MemoryStream())
        {
            using (FileStream fileStream = File.OpenRead(filePath))
            {
                fileStream.CopyTo(memoryStream);
            }

            // Reset stream position before detection
            memoryStream.Position = 0;

            // Detect file format information directly from the stream
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(memoryStream);

            // Report encryption status
            Console.WriteLine($"Is the workbook encrypted? {formatInfo.IsEncrypted}");

            // If the workbook is not encrypted, load it normally to demonstrate the Settings property
            if (!formatInfo.IsEncrypted)
            {
                memoryStream.Position = 0; // Reset again for loading
                Workbook workbook = new Workbook(memoryStream);
                Console.WriteLine($"Workbook.Settings.IsEncrypted: {workbook.Settings.IsEncrypted}");
            }
        }
    }
}
