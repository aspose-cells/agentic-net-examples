using Aspose.Cells;
using System;

class WriteProtectionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Apply write protection with a password
        wb.Settings.WriteProtection.Password = "ownerPwd";

        // Save the write‑protected workbook
        string protectedPath = "WriteProtected.xlsx";
        wb.Save(protectedPath);

        // Load the protected workbook
        Workbook loadedWb = new Workbook(protectedPath);

        // Check if the workbook is write protected
        Console.WriteLine("Is write protected: " + loadedWb.Settings.WriteProtection.IsWriteProtected);

        // Validate the password
        bool isValid = loadedWb.Settings.WriteProtection.ValidatePassword("ownerPwd");
        Console.WriteLine("Password valid: " + isValid);

        // Remove write protection by clearing the password
        if (isValid)
        {
            loadedWb.Settings.WriteProtection.Password = null; // or string.Empty
        }

        // Save the workbook without write protection
        string unprotectedPath = "WriteUnprotected.xlsx";
        loadedWb.Save(unprotectedPath);

        // Verify that protection has been removed
        Workbook finalWb = new Workbook(unprotectedPath);
        Console.WriteLine("Is write protected after removal: " + finalWb.Settings.WriteProtection.IsWriteProtected);
    }
}