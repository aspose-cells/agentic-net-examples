// Title: Generate a secure random password, apply workbook/worksheet protection, and retrieve its hash with Aspose.Cells (C#)
// Description: C# example that creates a 12‑character cryptographically strong password using System.Security.Cryptography, sets it as Workbook.Settings.WriteProtection.Password and Worksheet.Protection.Password, protects the sheet, obtains the integer hash via GetPasswordHash(), prints the password and hash, and saves the workbook as SecurePasswordDemo.xlsx.
// Keywords: Aspose.Cells random password generation | C# Excel write protection password | worksheet protection hash Aspose.Cells | cryptographic password for Excel | GetPasswordHash example | secure Excel file protection C# | Aspose.Cells password hash retrieval | generate secure password System.Security.Cryptography
// Common Searches: How to set a random write‑protection password for an Excel workbook using Aspose.Cells | Retrieve worksheet protection password hash with Aspose.Cells C# | Generate cryptographically secure password for Excel file protection | Aspose.Cells example for workbook and sheet protection | C# code to get password hash of protected worksheet
// Developer Intent: Create a strong random password, apply it to workbook and worksheet protection, obtain its hash, and save the protected Excel file.
// Use Cases: Automate generation of uniquely password‑protected reports where each file receives a runtime‑generated password. | Log the password hash for compliance auditing while keeping the actual password undisclosed. | Validate user‑entered passwords against stored hashes before allowing edits to protected worksheets.
// AI Prompts: Write C# code that uses Aspose.Cells to generate a cryptographically secure password, set it as the workbook write‑protection password, protect a worksheet, and output the password hash. | Show how to retrieve and display the integer hash of a worksheet protection password using Aspose.Cells GetPasswordHash(). | Explain how to compare a user‑supplied password with the hash returned by sheet.Protection.GetPasswordHash() in an Aspose.Cells workflow.

using System;
using System.Security.Cryptography;
using Aspose.Cells;

// C# example that creates a 12‑character cryptographically strong password using System.Security.Cryptography, sets it as Workbook.Settings.WriteProtection.Password and Worksheet.Protection.Password, protects the sheet, obtains the integer hash via GetPasswordHash(), prints the password and hash, and saves the workbook as SecurePasswordDemo.xlsx.
class SecurePasswordDemo
{
    static void Main()
    {
        // Generate a cryptographically strong random password
        string password = GenerateSecurePassword(12);
        Console.WriteLine($"Generated password: {password}");

        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set the write‑protection password (password to modify the file)
        workbook.Settings.WriteProtection.Password = password;

        // Protect the worksheet with the same password to obtain its hash
        sheet.Protection.Password = password;
        sheet.Protect(ProtectionType.All);

        // Retrieve and log the hash of the worksheet protection password
        int passwordHash = sheet.Protection.GetPasswordHash();
        Console.WriteLine($"Password hash: {passwordHash}");

        // Save the workbook
        workbook.Save("SecurePasswordDemo.xlsx");
    }

    // Generates a random alphanumeric password of the specified length
    private static string GenerateSecurePassword(int length)
    {
        const string allowedChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+";
        byte[] randomBytes = new byte[length];
        using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(randomBytes);
        }

        char[] chars = new char[length];
        for (int i = 0; i < length; i++)
        {
            chars[i] = allowedChars[randomBytes[i] % allowedChars.Length];
        }
        return new string(chars);
    }
}
