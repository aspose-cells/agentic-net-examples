// Title: Load Excel from UNC share, detect encryption, and manage permission errors using Aspose.Cells for .NET
// Description: Shows how to use Aspose.Cells in C# to detect the format and encryption of an Excel file on a UNC network share, load it with a password only when needed, and capture UnauthorizedAccessException, IOException, and other errors for logging.
// Keywords: Aspose.Cells | C# load Excel UNC path | detect encrypted workbook | FileFormatUtil | LoadOptions password | UnauthorizedAccessException handling | network share Excel | Excel encryption detection | Aspose.Cells exception handling | UNC network share access
// Common Searches: Aspose.Cells detect encrypted Excel file on network share | C# load Excel from UNC path with password | How to catch UnauthorizedAccessException in Aspose.Cells | FileFormatUtil DetectFileFormat example | LoadOptions password for encrypted workbook Aspose.Cells
// Developer Intent: Load a workbook from a network location, determine if it is password‑protected, and log any access‑permission or I/O problems.
// Use Cases: Check encryption before opening to avoid unnecessary password prompts | Read Excel files stored on shared drives in enterprise environments | Provide clear error logs for insufficient share permissions or missing files | Integrate secure password retrieval before loading encrypted workbooks
// AI Prompts: Generate C# code that uses Aspose.Cells to open an Excel file from a UNC path, detect encryption with FileFormatUtil, and load it with a password via LoadOptions only when required, including handling for UnauthorizedAccessException and IOException. | Explain how to extract detailed permission‑error information from UnauthorizedAccessException when Aspose.Cells accesses a workbook on a network share. | Show how to replace the hard‑coded password with a call to Azure Key Vault or Windows Credential Manager in the Aspose.Cells loading example. | Provide a logging strategy (e.g., using Serilog) for capturing I/O and security exceptions while loading workbooks from shared folders.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkExample
{
    // Shows how to use Aspose.Cells in C# to detect the format and encryption of an Excel file on a UNC network share, load it with a password only when needed, and capture UnauthorizedAccessException, IOException, and other errors for logging.
    class Program
    {
        static void Main()
        {
            // Path to the workbook on a network share
            string networkFilePath = @"\\ServerName\ShareFolder\SampleWorkbook.xlsx";

            try
            {
                // Detect file format and encryption status without opening the workbook
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(networkFilePath);
                Console.WriteLine($"File detected. Encrypted: {formatInfo.IsEncrypted}");

                // Load the workbook, providing a password only if the file is encrypted
                Workbook workbook;
                if (formatInfo.IsEncrypted)
                {
                    // If you know the password, set it here; otherwise loading will fail
                    var loadOptions = new LoadOptions { Password = "YourPasswordIfKnown" };
                    workbook = new Workbook(networkFilePath, loadOptions);
                }
                else
                {
                    workbook = new Workbook(networkFilePath);
                }

                Console.WriteLine("Workbook loaded successfully.");
                // Additional processing can be done here, e.g., accessing worksheets
            }
            catch (UnauthorizedAccessException uaEx)
            {
                // Log permission related issues
                Console.WriteLine($"Access permission error: {uaEx.Message}");
            }
            catch (IOException ioEx)
            {
                // Log I/O errors such as file not found or network problems
                Console.WriteLine($"I/O error while accessing the file: {ioEx.Message}");
            }
            catch (Exception ex)
            {
                // Log any other unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
