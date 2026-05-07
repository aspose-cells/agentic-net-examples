using System;
using Aspose.Cells;

class ValidateWriteProtection
{
    static void Main()
    {
        // Path to the workbook that may have a modification (write) password
        string filePath = "protected.xlsx";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(filePath);

        // Access the write protection settings of the workbook
        WriteProtection writeProtection = workbook.Settings.WriteProtection;

        // Determine if the workbook is write‑protected
        if (writeProtection.IsWriteProtected)
        {
            // The password you want to verify
            string passwordToValidate = "owner";

            // Use the ValidatePassword method to check the password
            bool isPasswordCorrect = writeProtection.ValidatePassword(passwordToValidate);

            Console.WriteLine($"Password '{passwordToValidate}' is valid: {isPasswordCorrect}");
        }
        else
        {
            Console.WriteLine("The workbook is not write protected.");
        }

        // Clean up resources
        workbook.Dispose();
    }
}