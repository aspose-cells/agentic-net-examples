// Title: C# – Verify Excel Workbook Is Encrypted and Password‑Protected with Aspose.Cells
// Description: A C# utility that loads an Excel file using Aspose.Cells LoadOptions, checks Workbook.Settings.IsEncrypted and Workbook.IsWorkbookProtectedWithPassword, and returns true only when the workbook is both encrypted and password‑protected. Includes file‑existence validation and exception handling.
// Keywords: Aspose.Cells C# encryption check | Workbook.IsWorkbookProtectedWithPassword example | LoadOptions password Excel .NET | detect encrypted Excel file | verify workbook protection status | C# Excel security validation | Aspose.Cells encrypted workbook detection
// Common Searches: how to check if Excel file is encrypted with Aspose.Cells | Aspose.Cells determine workbook password protection | C# load encrypted workbook using password | IsWorkbookProtectedWithPassword usage | verify both encryption and password protection in Excel .NET
// Developer Intent: Determine whether a given Excel workbook is simultaneously encrypted and protected by a password using Aspose.Cells for .NET.
// Use Cases: Screen uploaded spreadsheets in a web service to ensure they meet corporate encryption policies. | Automate acceptance criteria for secure data pipelines that require both encryption and password protection. | Log or reject Excel files that lack either encryption or password protection before further processing.
// AI Prompts: Generate unit tests for IsEncryptedAndPasswordProtected covering correct password, wrong password, and unprotected files. | Extend the method to return an enum indicating none, encrypted only, password only, or both. | Create detailed error messages that differentiate missing file, invalid password, and corrupted workbook scenarios.

using System;
using System.IO;
using Aspose.Cells;

// A C# utility that loads an Excel file using Aspose.Cells LoadOptions, checks Workbook.Settings.IsEncrypted and Workbook.IsWorkbookProtectedWithPassword, and returns true only when the workbook is both encrypted and password‑protected. Includes file‑existence validation and exception handling.
public class WorkbookProtectionChecker
{
    /// <param name="filePath">Full path to the workbook file.</param>
    /// <param name="openPassword">Password used to open the workbook (if it is encrypted).</param>
    /// <returns>True if the workbook is encrypted and also protected with a password; otherwise false.</returns>
    public static bool IsEncryptedAndPasswordProtected(string filePath, string openPassword)
    {
        // Verify that the file exists to avoid FileNotFoundException.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return false;
        }

        try
        {
            // Load the workbook using the provided password (if any).
            LoadOptions loadOptions = new LoadOptions
            {
                Password = openPassword
            };

            Workbook workbook = new Workbook(filePath, loadOptions);

            // Check encryption status.
            bool isEncrypted = workbook.Settings.IsEncrypted;

            // Check workbook protection status.
            bool isProtectedWithPassword = workbook.IsWorkbookProtectedWithPassword;

            return isEncrypted && isProtectedWithPassword;
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or checking.
            Console.WriteLine($"Error processing workbook: {ex.Message}");
            return false;
        }
    }

    // Example usage
    public static void Main()
    {
        string path = "protected_encrypted.xlsx";
        string password = "myPassword";

        bool result = IsEncryptedAndPasswordProtected(path, password);
        Console.WriteLine($"Workbook is both encrypted and password protected: {result}");
    }
}
