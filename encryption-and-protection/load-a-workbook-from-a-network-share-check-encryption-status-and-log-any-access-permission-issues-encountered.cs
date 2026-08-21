// Title: Load Excel workbook from UNC share, detect encryption, handle permission errors – Aspose.Cells .NET
// Description: C# example that uses Aspose.Cells to detect the format and encryption state of an Excel file on a network (UNC) share with FileFormatUtil.DetectFileFormat, loads the workbook when possible, checks Workbook.Settings.IsEncrypted, and logs UnauthorizedAccessException, IOException, or other errors.
// Keywords: Aspose.Cells UNC path | load workbook from network share | detect encrypted Excel file | FileFormatUtil DetectFileFormat | Workbook.Settings.IsEncrypted | handle UnauthorizedAccessException | C# Excel file permission error | network share Excel access Aspose
// Common Searches: Aspose.Cells load Excel from UNC path | check if Excel file on network share is password protected | detect encryption without opening workbook Aspose.Cells | log permission denied error when opening Excel file | C# example for FileFormatUtil encryption detection
// Developer Intent: Load an Excel workbook located on a network share, determine whether it is encrypted, and capture any access‑permission or I/O problems.
// Use Cases: Validate encryption status before opening a shared workbook to avoid unexpected password prompts. | Record detailed logs when a user lacks read rights on a UNC folder for compliance auditing. | Use lightweight format detection to skip loading large encrypted files in batch processes.
// AI Prompts: Write C# code that opens an Excel file from a UNC path with Aspose.Cells, checks Workbook.Settings.IsEncrypted, and logs UnauthorizedAccessException with full details. | Show how to use FileFormatUtil.DetectFileFormat to identify a password‑protected workbook on a network share before calling new Workbook(). | Provide best‑practice guidelines for handling I/O and permission exceptions when accessing Excel files on remote shares using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkExample
{
    // C# example that uses Aspose.Cells to detect the format and encryption state of an Excel file on a network (UNC) share with FileFormatUtil.DetectFileFormat, loads the workbook when possible, checks Workbook.Settings.IsEncrypted, and logs UnauthorizedAccessException, IOException, or other errors.
    class Program
    {
        static void Main()
        {
            // Path to the workbook on a network share
            string networkFilePath = @"\\server\share\folder\example.xlsx";

            try
            {
                // Detect file format and encryption status without opening the workbook
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(networkFilePath);
                Console.WriteLine($"File detected as {formatInfo.FileFormatType}");
                Console.WriteLine($"IsEncrypted (detected): {formatInfo.IsEncrypted}");

                // Load the workbook (no password supplied; will fail if encrypted)
                Workbook workbook = new Workbook(networkFilePath);
                // After loading, also check the workbook settings for encryption
                Console.WriteLine($"Workbook.Settings.IsEncrypted: {workbook.Settings.IsEncrypted}");
            }
            catch (UnauthorizedAccessException ex)
            {
                // Log permission issues when accessing the network share
                Console.WriteLine($"Access denied to '{networkFilePath}'. Details: {ex.Message}");
            }
            catch (IOException ex)
            {
                // Log other I/O related problems (e.g., file not found, network errors)
                Console.WriteLine($"I/O error while accessing '{networkFilePath}'. Details: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General exception handling for unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
