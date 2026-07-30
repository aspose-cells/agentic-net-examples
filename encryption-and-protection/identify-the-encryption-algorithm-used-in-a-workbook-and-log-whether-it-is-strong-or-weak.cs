// Title: C# – Detect Excel Workbook Encryption Algorithm and Strength with Aspose.Cells
// Description: Uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify if an Excel file is encrypted, determines the algorithm (SHA‑AES for modern OOXML formats or legacy for older binaries), classifies the strength as Strong or Weak, and logs the result to the console.
// Keywords: Aspose.Cells | C# encryption detection | Excel workbook encryption algorithm | SHA‑AES Excel | legacy Excel encryption | FileFormatUtil | DetectFileFormat | .NET | strong vs weak encryption | Excel security compliance
// Common Searches: Aspose.Cells detect encrypted Excel file C# | How to check Excel encryption algorithm with Aspose.Cells | Identify SHA‑AES encryption in XLSX using .NET | Determine if Excel workbook uses weak legacy encryption | C# code to log Excel encryption strength
// Developer Intent: Identify the encryption algorithm of an Excel workbook and report whether it is strong or weak.
// Use Cases: Automated security scan of uploaded Excel files to enforce strong encryption policies | Audit logging of workbook encryption type for regulatory compliance | Conditional processing: reject files encrypted with legacy weak algorithms in data‑import services
// AI Prompts: Generate C# code that uses Aspose.Cells to detect if a workbook is encrypted, returns the algorithm name, and indicates 'Strong' for SHA‑AES and 'Weak' for legacy formats. | Create a reusable function returning an enum (NotEncrypted, Strong, Weak) based on FileFormatUtil.DetectFileFormat for any Excel file. | Provide robust error handling for missing, corrupted, or permission‑restricted Excel files when checking encryption status with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    // Uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify if an Excel file is encrypted, determines the algorithm (SHA‑AES for modern OOXML formats or legacy for older binaries), classifies the strength as Strong or Weak, and logs the result to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be examined
            string workbookPath = "sample.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
                return;
            }

            try
            {
                // Detect file format and encryption status
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(workbookPath);

                bool isEncrypted = formatInfo.IsEncrypted;
                Console.WriteLine($"Workbook encrypted: {isEncrypted}");

                // Determine encryption algorithm and strength
                string algorithm = "None";
                string strength = "N/A";

                if (isEncrypted)
                {
                    // Modern OOXML formats (XLSX, XLSB, XLSM, etc.) use SHA‑AES encryption (strong)
                    if (formatInfo.FileFormatType == FileFormatType.Xlsx ||
                        formatInfo.FileFormatType == FileFormatType.Xlsb ||
                        formatInfo.FileFormatType == FileFormatType.Xlsm)
                    {
                        algorithm = "SHA‑AES (Office 2007+)";
                        strength = "Strong";
                    }
                    // For any other encrypted format (e.g., legacy binary XLS) treat as weak/unknown
                    else
                    {
                        algorithm = "Legacy/Unknown (Office 97‑2003 or other)";
                        strength = "Weak or Unknown";
                    }
                }

                Console.WriteLine($"Encryption algorithm: {algorithm}");
                Console.WriteLine($"Encryption strength: {strength}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors (e.g., corrupted file, permission issues)
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
