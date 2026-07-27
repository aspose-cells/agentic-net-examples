// Title: Validate Write‑Protection Password of an Excel Workbook Using Aspose.Cells for .NET
// Description: Load an Excel file in read‑only mode, access its Settings.WriteProtection, and call WriteProtection.ValidatePassword to determine if a supplied password unlocks the workbook, returning true or false.
// Keywords: Aspose.Cells | C# validate Excel password | write protection password | Workbook.ValidatePassword | Aspose.Cells WriteProtection | Excel file security .NET | check workbook password without editing
// Common Searches: Aspose.Cells validate workbook password C# | How to check write‑protected Excel file password .NET | Validate Excel write protection without opening file | C# code to test Excel password using Aspose | Determine if Excel workbook is write‑protected Aspose.Cells
// Developer Intent: Check whether a given password matches the write‑protection password of an Excel workbook.
// Use Cases: Prevent unauthorized edits by verifying the password before allowing modifications. | Skip or flag files in batch jobs when the write‑protection password cannot be confirmed. | Provide a service endpoint that reports editability of uploaded Excel workbooks.
// AI Prompts: Create a C# function that uses Aspose.Cells to verify a write‑protected Excel workbook password and returns a boolean. | Write unit tests for the ValidatePassword method covering correct, incorrect, and unprotected workbook scenarios. | Explain the internal hashing mechanism used by Aspose.Cells WriteProtection.ValidatePassword.

using System;
using System.IO;
using Aspose.Cells;

// Load an Excel file in read‑only mode, access its Settings.WriteProtection, and call WriteProtection.ValidatePassword to determine if a supplied password unlocks the workbook, returning true or false.
public class WorkbookWriteProtectionValidator
{
    /// <param name="filePath">Full path to the workbook file.</param>
    /// <param name="password">Password to validate.</param>
    /// <returns>True if the password is correct; otherwise false.</returns>
    public static bool ValidatePassword(string filePath, string password)
    {
        // Load the workbook (read‑only, no need to save later)
        Workbook workbook = new Workbook(filePath);

        // Access the write‑protection settings of the workbook
        WriteProtection writeProtection = workbook.Settings.WriteProtection;

        // If the workbook is not write‑protected, any password is considered invalid
        if (!writeProtection.IsWriteProtected)
            return false;

        // Use the built‑in ValidatePassword method to compare the supplied password
        return writeProtection.ValidatePassword(password);
    }

    // Example usage
    public static void Main()
    {
        string path = "WriteProtectedWorkbook.xlsx";
        string correctPassword = "owner";
        string wrongPassword = "wrong";

        bool isCorrectValid = ValidatePassword(path, correctPassword);
        Console.WriteLine($"Password '{correctPassword}' is valid: {isCorrectValid}");

        bool isWrongValid = ValidatePassword(path, wrongPassword);
        Console.WriteLine($"Password '{wrongPassword}' is valid: {isWrongValid}");
    }
}
