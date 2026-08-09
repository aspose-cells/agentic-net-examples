// Title: Detect Encryption in an XLS Workbook and Retrieve Encryption Details with Aspose.Cells for .NET
// Description: Demonstrates how to use Aspose.Cells' FileFormatUtil.DetectFileFormat to check if an XLS file is encrypted, read the IsEncrypted flag, and understand that the specific encryption algorithm is not exposed via the API. Includes guidance for attempting a password‑based load after detection.
// Keywords: Aspose.Cells | C# | XLS encryption detection | FileFormatUtil DetectFileFormat | IsEncrypted flag | Excel password protection | encryption algorithm | detect encrypted workbook | Aspose.Cells encryption info | Excel file security .NET
// Common Searches: How to check if an XLS file is encrypted using Aspose.Cells | Aspose.Cells get encryption status of Excel workbook | Retrieve encryption algorithm of XLS with Aspose.Cells | Detect password protection on Excel file C# | Aspose.Cells FileFormatUtil encrypted file detection
// Developer Intent: Determine whether an XLS workbook is encrypted and learn what encryption information Aspose.Cells can provide.
// Use Cases: Screen user‑uploaded XLS files for password protection before further processing. | Log encryption status of incoming Excel files for compliance and audit trails. | After detecting encryption, attempt to open the workbook with a known password. | Provide clear feedback to end‑users about required file security settings.
// AI Prompts: Generate C# code using Aspose.Cells that returns a bool indicating encryption and, if true, tries to open the workbook with a supplied password. | Explain why Aspose.Cells does not expose the exact encryption algorithm for legacy XLS files and suggest alternative ways to infer the algorithm. | Create a helper method that returns encryption status and a descriptive message, handling cases where algorithm details are unavailable.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells' FileFormatUtil.DetectFileFormat to check if an XLS file is encrypted, read the IsEncrypted flag, and understand that the specific encryption algorithm is not exposed via the API. Includes guidance for attempting a password‑based load after detection.
class VerifyEncryption
{
    static void Main()
    {
        // Path to the XLS file to be examined
        string filePath = "sample.xls";

        // Detect file format and retrieve encryption information
        FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

        // Output whether the file is encrypted
        Console.WriteLine($"Is the file encrypted? {fileInfo.IsEncrypted}");

        // Aspose.Cells does not expose the specific encryption algorithm directly.
        // If needed, you can attempt to load the workbook with a password to verify access.
        if (fileInfo.IsEncrypted)
        {
            Console.WriteLine("Encryption algorithm information is not directly available via the API.");
        }
    }
}
