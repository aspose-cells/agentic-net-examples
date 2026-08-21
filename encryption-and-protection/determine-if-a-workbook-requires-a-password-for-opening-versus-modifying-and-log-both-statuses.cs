// Title: Check if an Excel workbook requires a password to open or to edit using Aspose.Cells for .NET
// Description: C# example that uses Aspose.Cells to detect whether an Excel file is encrypted (password needed to open) via FileFormatUtil.DetectFileFormat, then, if not encrypted, loads the workbook to read WriteProtection.IsWriteProtected and IsWorkbookProtectedWithPassword. The three protection states are written to the console.
// Keywords: Aspose.Cells detect encrypted workbook | Excel file password open .NET | check write protection Aspose.Cells | workbook structure protection C# | FileFormatUtil DetectFileFormat | Workbook.Settings.WriteProtection | IsWorkbookProtectedWithPassword | Excel security audit Aspose
// Common Searches: how to know if an Excel file is password protected for opening using Aspose.Cells | Aspose.Cells .NET check if workbook is write‑protected | detect workbook encryption and write protection with Aspose.Cells | C# get Excel file protection status Aspose | Aspose.Cells determine if workbook structure is password protected
// Developer Intent: Identify whether an Excel workbook is encrypted (requires a password to open) and whether it is write‑protected or structure‑protected, then output those flags.
// Use Cases: Validate protection level before extracting or modifying data. | Log encryption and write‑protection status for compliance audits. | Conditionally load a workbook with a known password only when it is encrypted.
// AI Prompts: Generate C# code that opens an encrypted Excel file with a supplied password using Aspose.Cells, then reports write protection and structure protection. | Provide error‑handling patterns for cases where the workbook is encrypted but the opening password is missing, while still returning the encryption flag. | Show how to embed the protection‑status checks into a processing pipeline that skips write‑protected workbooks.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordCheck
{
    // C# example that uses Aspose.Cells to detect whether an Excel file is encrypted (password needed to open) via FileFormatUtil.DetectFileFormat, then, if not encrypted, loads the workbook to read WriteProtection.IsWriteProtected and IsWorkbookProtectedWithPassword. The three protection states are written to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be examined
            string workbookPath = "sample.xlsx";

            // Detect if the file is encrypted (requires a password to open)
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(workbookPath);
            bool isEncrypted = formatInfo.IsEncrypted;

            // Variables to hold modification protection status
            bool isWriteProtected = false;
            bool isWorkbookProtectedWithPassword = false;

            // If the workbook is not encrypted, we can load it to inspect further protection settings
            if (!isEncrypted)
            {
                // Load the workbook (no password needed)
                Workbook workbook = new Workbook(workbookPath);

                // Check if the workbook is write‑protected (requires a password to modify)
                isWriteProtected = workbook.Settings.WriteProtection.IsWriteProtected;

                // Check if the workbook structure or window is protected with a password
                isWorkbookProtectedWithPassword = workbook.IsWorkbookProtectedWithPassword;
            }
            else
            {
                // When encrypted, additional protection details cannot be read without the opening password.
                // Optionally, load with a known password here to inspect further, if available.
            }

            // Log the results
            Console.WriteLine($"Requires password to open (IsEncrypted): {isEncrypted}");
            Console.WriteLine($"Requires password to modify (WriteProtection.IsWriteProtected): {isWriteProtected}");
            Console.WriteLine($"Workbook structure/window protected with password (IsWorkbookProtectedWithPassword): {isWorkbookProtectedWithPassword}");
        }
    }
}
