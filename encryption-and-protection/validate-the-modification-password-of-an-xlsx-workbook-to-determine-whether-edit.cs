using System;
using Aspose.Cells;

class ValidateWriteProtection
{
    static void Main()
    {
        // Path to the workbook that may have write‑protection enabled
        string filePath = "protected.xlsx";

        // Load the workbook (write‑protection does not require a password to open)
        Workbook workbook = new Workbook(filePath);

        // Access the write‑protection settings of the workbook
        WriteProtection writeProtection = workbook.Settings.WriteProtection;

        // Indicate whether the workbook is write‑protected
        Console.WriteLine("Is workbook write protected? " + writeProtection.IsWriteProtected);

        // Password to test – replace with the actual password you expect
        string passwordToTest = "owner";

        // Validate the password against the write‑protection password
        bool isValid = writeProtection.ValidatePassword(passwordToTest);
        Console.WriteLine($"Password '{passwordToTest}' is valid: {isValid}");

        // Demonstrate validation with an incorrect password
        string wrongPassword = "wrong";
        bool isWrongValid = writeProtection.ValidatePassword(wrongPassword);
        Console.WriteLine($"Password '{wrongPassword}' is valid: {isWrongValid}");
    }
}