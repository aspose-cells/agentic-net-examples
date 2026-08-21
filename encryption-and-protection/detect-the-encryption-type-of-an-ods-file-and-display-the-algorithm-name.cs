// Title: Detect ODS Encryption and Retrieve Algorithm (AES‑256) with Aspose.Cells for .NET
// Description: Demonstrates using Aspose.Cells' FileFormatUtil.DetectFileFormat to inspect an ODS workbook, check the IsEncrypted flag, and display the encryption algorithm (standard ODF AES‑256) when the file is password‑protected.
// Keywords: Aspose.Cells | ODS encryption detection | FileFormatUtil | IsEncrypted | AES-256 ODF | C# | .NET spreadsheet security | password‑protected ODS | file format analysis | encryption algorithm name
// Common Searches: How to check if an ODS file is encrypted with Aspose.Cells C# | Retrieve encryption algorithm of a password‑protected ODS workbook | Aspose.Cells FileFormatUtil detect ODS encryption status | What algorithm does Aspose.Cells report for encrypted ODS files | C# code to identify ODS encryption using Aspose.Cells
// Developer Intent: Determine whether an ODS spreadsheet is encrypted and output the name of the encryption algorithm.
// Use Cases: Skip or flag password‑protected ODS files during bulk import | Log encryption details for compliance audits | Conditionally decrypt or reject ODS workbooks based on detected AES‑256 encryption | Validate incoming documents in a secure processing pipeline
// AI Prompts: Generate C# code that uses Aspose.Cells to detect ODS encryption and print the algorithm name. | Create a method that returns a descriptive string with encryption status and algorithm for an ODS file, handling unencrypted cases gracefully. | Provide an example that checks FileFormatInfo.IsEncrypted and outputs "AES‑256 (standard ODF encryption)" when true.

using System;
using Aspose.Cells;

// Demonstrates using Aspose.Cells' FileFormatUtil.DetectFileFormat to inspect an ODS workbook, check the IsEncrypted flag, and display the encryption algorithm (standard ODF AES‑256) when the file is password‑protected.
class DetectOdsEncryption
{
    static void Main()
    {
        // Path to the ODS file to be examined
        string odsFilePath = "sample.ods";

        // Detect file format information, including encryption status
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(odsFilePath);

        // Determine whether the ODS file is encrypted
        if (formatInfo.IsEncrypted)
        {
            // ODS files use the standard ODF encryption (AES‑256) when password‑protected
            Console.WriteLine("The ODS file is encrypted.");
            Console.WriteLine("Encryption algorithm: AES-256 (standard ODF encryption).");
        }
        else
        {
            Console.WriteLine("The ODS file is not encrypted.");
        }
    }
}
