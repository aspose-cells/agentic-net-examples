// Title: Detect encrypted Excel workbook and validate password complexity using Aspose.Cells for .NET
// Common Searches: how to determine if an Excel file is password protected with Aspose.Cells in C# | C# code to open an encrypted XLSX using Aspose.Cells LoadOptions password | validate Excel password meets corporate complexity rules after loading with Aspose.Cells | check workbook encryption status before applying password policy in .NET | Aspose.Cells exception handling for encrypted workbook detection
// Tags: Aspose.Cells encrypted workbook detection | LoadOptions password for encrypted XLSX | C# password complexity validation Excel | Workbook encryption status Aspose.Cells | Corporate Excel password policy .NET

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example shows how to detect whether an XLSX workbook is password‑protected by attempting to load it without credentials and catching the resulting exception. If encrypted, it reloads the file using LoadOptions with the supplied password, then evaluates the password against typical organizational complexity requirements (minimum length, uppercase, lowercase, digit, special character) and prints both the encryption status and the complexity result.
class WorkbookEncryptionChecker
{
    static void Main()
    {
        // Path to the workbook to be examined
        string workbookPath = "encrypted.xlsx";

        // Password supplied for opening the workbook (could be obtained from user input)
        string password = "P@ssw0rd!";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"File not found: {workbookPath}");
            return;
        }

        bool isEncrypted = false;
        Workbook workbook = null;

        // Attempt to load without a password to determine if the workbook is encrypted
        try
        {
            workbook = new Workbook(workbookPath);
            isEncrypted = false;
        }
        catch (CellsException)
        {
            // If an exception occurs, assume the workbook is encrypted
            isEncrypted = true;
        }

        // If encrypted, load it using the supplied password
        if (isEncrypted)
        {
            try
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                workbook = new Workbook(workbookPath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to open workbook with the provided password: {ex.Message}");
                return;
            }
        }

        // Validate password against organizational complexity rules:
        // - Minimum 8 characters
        // - At least one uppercase letter
        // - At least one lowercase letter
        // - At least one digit
        // - At least one special character
        bool meetsComplexity = false;
        if (!string.IsNullOrEmpty(password))
        {
            bool hasUpper = password.Any(char.IsUpper);
            bool hasLower = password.Any(char.IsLower);
            bool hasDigit = password.Any(char.IsDigit);
            bool hasSpecial = password.Any(ch => !char.IsLetterOrDigit(ch));
            meetsComplexity = password.Length >= 8 && hasUpper && hasLower && hasDigit && hasSpecial;
        }

        // Output the results
        Console.WriteLine($"Workbook Encrypted: {isEncrypted}");
        Console.WriteLine($"Password Meets Complexity: {meetsComplexity}");
    }
}
