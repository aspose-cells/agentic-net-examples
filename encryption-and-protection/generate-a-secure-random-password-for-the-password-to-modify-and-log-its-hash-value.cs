using System;
using System.Security.Cryptography;
using Aspose.Cells;

class SecurePasswordDemo
{
    static void Main()
    {
        // Generate a cryptographically strong random password (16 bytes, Base64 encoded)
        byte[] pwdBytes = new byte[16];
        RandomNumberGenerator.Fill(pwdBytes);
        string password = Convert.ToBase64String(pwdBytes);

        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the write‑protection password (password required to modify the file)
        workbook.Settings.WriteProtection.Password = password;

        // Apply the same password to worksheet protection so we can obtain its hash
        worksheet.Protection.Password = password;
        worksheet.Protect(ProtectionType.All);

        // Retrieve the hash of the worksheet protection password
        int passwordHash = worksheet.Protection.GetPasswordHash();

        // Log the generated password and its hash value
        Console.WriteLine($"Generated password: {password}");
        Console.WriteLine($"Password hash: {passwordHash}");

        // Save the workbook
        workbook.Save("SecurePasswordDemo.xlsx");
    }
}