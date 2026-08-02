// Title: Check if an Excel workbook is encrypted with Aspose.Cells for .NET (C#)
// Description: Loads a workbook from a file path, verifies the file exists, and uses Workbook.Settings.IsEncrypted to determine whether the Excel file is password‑protected. The sample prints the encryption flag and gracefully handles missing files or load errors.
// Keywords: Aspose.Cells | C# workbook encryption check | Workbook.Settings.IsEncrypted | detect password-protected Excel | load encrypted Excel .NET | Excel file encryption status | Aspose.Cells .NET example | Excel security verification | C# file existence check | Aspose.Cells API
// Common Searches: Aspose.Cells how to know if Excel file is password protected | C# check Excel encryption Aspose | Workbook.Settings.IsEncrypted example | detect encrypted workbook before opening | Aspose.Cells .NET encryption status | C# verify Excel file requires password | load Excel file with Aspose and check encryption
// Developer Intent: Identify whether a given Excel workbook is encrypted (password‑protected) using Aspose.Cells for .NET.
// Use Cases: Skip processing of encrypted files to prevent runtime exceptions | Prompt users for a password only when the workbook is encrypted | Generate a report of encryption status across multiple workbooks | Integrate an encryption check into automated import pipelines | Log security‑compliance information for Excel documents
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, checks Workbook.Settings.IsEncrypted, and if true asks the user for a password before reopening. | Create a C# console app that scans a folder of .xlsx files and outputs each file's encryption status using Aspose.Cells. | Explain how to catch the exception thrown when opening an encrypted workbook without a password in Aspose.Cells and handle it gracefully.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    // Loads a workbook from a file path, verifies the file exists, and uses Workbook.Settings.IsEncrypted to determine whether the Excel file is password‑protected. The sample prints the encryption flag and gracefully handles missing files or load errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file on disk
            string filePath = @"C:\Path\To\YourWorkbook.xlsx";

            // Verify that the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(filePath);

                // Check whether the workbook is encrypted (requires a password to open)
                bool isEncrypted = workbook.Settings.IsEncrypted;

                // Output the result
                Console.WriteLine($"Workbook encrypted: {isEncrypted}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
