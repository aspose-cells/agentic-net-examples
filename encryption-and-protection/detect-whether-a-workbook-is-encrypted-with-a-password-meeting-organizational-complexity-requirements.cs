using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

class WorkbookEncryptionChecker
{
    // Checks password against typical organizational complexity rules:
    // - Minimum 8 characters
    // - At least one uppercase letter
    // - At least one lowercase letter
    // - At least one digit
    // - At least one special character
    static bool IsPasswordComplex(string password)
    {
        if (string.IsNullOrEmpty(password))
            return false;

        var pattern = @"^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[\W_]).{8,}$";
        return Regex.IsMatch(password, pattern);
    }

    // Detects encryption status and validates password complexity
    static void CheckWorkbook(string filePath, string password)
    {
        // Detect encryption without opening the workbook
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine($"File encrypted: {formatInfo.IsEncrypted}");

        if (!formatInfo.IsEncrypted)
        {
            Console.WriteLine("Workbook is not encrypted.");
            return;
        }

        // Verify password meets complexity requirements
        bool complexityOk = IsPasswordComplex(password);
        Console.WriteLine($"Password meets complexity requirements: {complexityOk}");

        // Attempt to load the workbook using the supplied password
        LoadOptions loadOptions = new LoadOptions { Password = password };
        try
        {
            Workbook wb = new Workbook(filePath, loadOptions);
            // After successful load, confirm the workbook reports as encrypted
            Console.WriteLine($"Workbook loaded successfully. Settings.IsEncrypted: {wb.Settings.IsEncrypted}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook with provided password: {ex.Message}");
        }
    }

    static void Main()
    {
        // Example usage
        string filePath = "encrypted.xlsx";
        string password = "P@ssw0rd!"; // Replace with the password to test
        CheckWorkbook(filePath, password);
    }
}