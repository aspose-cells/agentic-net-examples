using System;
using Aspose.Cells;

class CheckWriteProtectionPassword
{
    static void Main()
    {
        // Path to the Excel file that may have write‑protection enabled
        string filePath = "WriteProtectedWorkbook.xlsx";

        // Load the workbook (write‑protection does not require a password to open)
        Workbook workbook = new Workbook(filePath);

        // Access the write‑protection settings of the workbook
        WriteProtection writeProtection = workbook.Settings.WriteProtection;

        // Determine whether the workbook is write‑protected
        if (writeProtection.IsWriteProtected)
        {
            // The password we want to verify
            string passwordToCheck = "owner";

            // Validate the supplied password against the workbook's write‑protection password
            bool isPasswordValid = writeProtection.ValidatePassword(passwordToCheck);

            Console.WriteLine($"Password '{passwordToCheck}' is valid: {isPasswordValid}");
        }
        else
        {
            Console.WriteLine("The workbook is not write‑protected.");
        }
    }
}