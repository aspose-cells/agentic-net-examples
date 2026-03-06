using System;
using Aspose.Cells;

class CheckWriteProtection
{
    static void Main(string[] args)
    {
        // Path to the workbook and the password to test can be supplied via command‑line arguments.
        string filePath = args.Length > 0 ? args[0] : "protected.xlsx";
        string passwordToTest = args.Length > 1 ? args[1] : "owner";

        // Load the workbook (write‑protection does not require a password to open).
        Workbook workbook = new Workbook(filePath);

        // Get the write‑protection settings.
        WriteProtection writeProtection = workbook.Settings.WriteProtection;

        // Determine whether the workbook is write‑protected.
        if (writeProtection.IsWriteProtected)
        {
            // Validate the supplied password against the write‑protection password.
            bool isPasswordValid = writeProtection.ValidatePassword(passwordToTest);
            Console.WriteLine($"Workbook is write‑protected. Password valid: {isPasswordValid}");

            // If the password is correct, remove write protection so the file can be modified.
            if (isPasswordValid)
            {
                writeProtection.Password = null; // clear the password
                Console.WriteLine("Write protection removed.");

                // Save the unprotected workbook.
                workbook.Save("unprotected.xlsx");
                Console.WriteLine("Saved unprotected workbook as 'unprotected.xlsx'.");
            }
        }
        else
        {
            Console.WriteLine("Workbook is not write‑protected.");
        }
    }
}