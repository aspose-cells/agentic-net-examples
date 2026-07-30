// Title: Check if an Excel workbook is encrypted and password‑protected with Aspose.Cells for .NET
// Description: C# helper that loads an Excel file using Aspose.Cells LoadOptions with a supplied password, then evaluates Workbook.Settings.IsEncrypted and Workbook.IsWorkbookProtectedWithPassword. Returns true only when the file requires a password to open (encryption) and its structure is also protected, with graceful handling of missing files and runtime errors.
// Keywords: Aspose.Cells encryption detection | C# check Excel password protection | Workbook.IsWorkbookProtectedWithPassword | LoadOptions password Aspose.Cells | verify encrypted workbook .NET | Excel file security Aspose | detect encrypted and protected workbook
// Common Searches: how to determine if an Excel file is encrypted and workbook protected using Aspose.Cells | C# method to verify both encryption and password protection in an Excel workbook | Aspose.Cells check if workbook requires a password to open and has structure protection | detect encrypted Excel file with Aspose.Cells .NET | verify workbook protection status programmatically
// Developer Intent: Identify whether a specific Excel workbook is both encrypted (requires a password to open) and protected with a workbook‑level password.
// Use Cases: Validate uploaded Excel files in a web portal to enforce encryption and structure protection before processing. | Filter or flag workbooks in a batch import pipeline that lack either encryption or password protection for compliance reporting. | Log security status of Excel documents in an automated audit system, rejecting those that are only partially protected.
// AI Prompts: Generate unit tests for IsWorkbookEncryptedAndPasswordProtected covering encrypted‑only, protected‑only, both, and none scenarios. | Refactor the helper to return an enum (None, EncryptedOnly, ProtectedOnly, Both) with descriptive messages. | Show how to integrate the method into an ASP.NET Core file‑upload endpoint that rejects insecure Excel files and returns a clear error response.

using System;
using System.IO;
using Aspose.Cells;

// C# helper that loads an Excel file using Aspose.Cells LoadOptions with a supplied password, then evaluates Workbook.Settings.IsEncrypted and Workbook.IsWorkbookProtectedWithPassword. Returns true only when the file requires a password to open (encryption) and its structure is also protected, with graceful handling of missing files and runtime errors.
public static class WorkbookProtectionHelper
{
    /// <param name="filePath">Full path to the workbook file.</param>
    /// <param name="password">Password used to open the workbook (if it is encrypted).</param>
    /// <returns>True if the workbook is encrypted and also protected with a password; otherwise false.</returns>
    public static bool IsWorkbookEncryptedAndPasswordProtected(string filePath, string password)
    {
        // Verify that the file exists to avoid FileNotFoundException.
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Workbook file not found: {filePath}");

        try
        {
            // Load the workbook using the supplied password.
            // If the workbook is not encrypted, the password is simply ignored.
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            Workbook workbook = new Workbook(filePath, loadOptions);

            // Check if the workbook requires a password to open (encryption).
            bool isEncrypted = workbook.Settings.IsEncrypted;

            // Check if the workbook's structure or window is protected with a password.
            bool isProtectedWithPassword = workbook.IsWorkbookProtectedWithPassword;

            // Return true only when both conditions are satisfied.
            return isEncrypted && isProtectedWithPassword;
        }
        catch (Exception ex)
        {
            // Log or rethrow as needed; here we return false to indicate the check failed.
            Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
            return false;
        }
    }
}

public class Program
{
    // Entry point required for console application.
    public static void Main(string[] args)
    {
        // Example usage: provide workbook path and password as arguments or hard‑code for testing.
        string workbookPath = args.Length > 0 ? args[0] : "sample.xlsx";
        string password = args.Length > 1 ? args[1] : "myPassword";

        try
        {
            bool result = WorkbookProtectionHelper.IsWorkbookEncryptedAndPasswordProtected(workbookPath, password);
            Console.WriteLine($"Workbook encrypted and password protected: {result}");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.Error.WriteLine(fnfEx.Message);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
