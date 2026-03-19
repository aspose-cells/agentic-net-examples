using System;
using Aspose.Cells;

class VerifyModificationPassword
{
    // Verifies the write‑protection (modification) password of an existing XLSX file.
    public static void Run(string filePath, string password)
    {
        // Load the workbook (no opening password assumed)
        Workbook workbook = new Workbook(filePath);

        // Determine whether the workbook has write‑protection enabled
        bool isWriteProtected = workbook.Settings.WriteProtection.IsWriteProtected;
        Console.WriteLine($"Workbook write‑protected: {isWriteProtected}");

        if (isWriteProtected)
        {
            // Validate the supplied modification password
            bool isValid = workbook.Settings.WriteProtection.ValidatePassword(password);
            Console.WriteLine($"Password '{password}' is valid: {isValid}");
        }
        else
        {
            Console.WriteLine("Workbook is not write‑protected; no password verification required.");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <exe> <excelFilePath> <modificationPassword>");
            return;
        }

        string filePath = args[0];
        string password = args[1];

        VerifyModificationPassword.Run(filePath, password);
    }
}