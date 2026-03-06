using System;
using Aspose.Cells;

class WriteProtectionPasswordCheck
{
    static void Main()
    {
        // Path for the temporary workbook
        string filePath = "WriteProtectedWorkbook.xlsx";

        // ---------- Create ----------
        // Create a new workbook and set a write‑protection password
        Workbook workbook = new Workbook();
        workbook.Settings.WriteProtection.Password = "owner";
        // Save the workbook (write‑protected)
        workbook.Save(filePath);

        // ---------- Load ----------
        // Load the saved workbook
        Workbook loadedWorkbook = new Workbook(filePath);

        // ---------- Check password ----------
        // Password to verify
        string passwordToValidate = "owner";

        // Use WriteProtection.ValidatePassword to check if the password matches
        bool isPasswordValid = loadedWorkbook.Settings.WriteProtection.ValidatePassword(passwordToValidate);

        Console.WriteLine($"Password '{passwordToValidate}' is valid: {isPasswordValid}");

        // Example of an invalid password
        string wrongPassword = "wrong";
        bool isWrongPasswordValid = loadedWorkbook.Settings.WriteProtection.ValidatePassword(wrongPassword);
        Console.WriteLine($"Password '{wrongPassword}' is valid: {isWrongPasswordValid}");
    }
}