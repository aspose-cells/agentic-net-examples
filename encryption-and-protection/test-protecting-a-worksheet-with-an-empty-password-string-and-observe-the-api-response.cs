// Title: Protect an Aspose.Cells worksheet with a blank password and verify persistence
// Description: C# example that protects the first worksheet of a new workbook using an empty password, checks Worksheet.IsProtected and Worksheet.Protection.IsProtectedWithPassword flags, validates the empty password with VerifyPassword, saves the file, reloads it, and confirms that the protection state remains unchanged.
// Keywords: Aspose.Cells protect worksheet empty password | Worksheet.Protect blank password C# | Worksheet.IsProtected flag | Worksheet.Protection.IsProtectedWithPassword false | VerifyPassword empty string Aspose.Cells | save protected worksheet Aspose.Cells | load worksheet protection state | C# Aspose.Cells protection API
// Common Searches: Aspose.Cells protect sheet with no password | IsProtected vs IsProtectedWithPassword Aspose.Cells | How to verify empty password on protected worksheet | Does blank password allow unprotecting in Aspose.Cells | Persist worksheet protection after saving Aspose.Cells
// Developer Intent: The developer wants to apply worksheet protection without a password, observe the resulting protection flags, and ensure the protection persists after the workbook is saved and reopened.
// Use Cases: Enable UI indicators for protected sheets while allowing users to unprotect without entering a password. | Test that a workbook saved with a blank password retains the same protection state when opened later. | Use IsProtected and IsProtectedWithPassword to decide whether to prompt for a password in a spreadsheet application.
// AI Prompts: Generate C# code that protects an Aspose.Cells worksheet with an empty password, prints protection flags, saves the workbook, reloads it, and re‑checks the flags. | Explain the difference between Worksheet.IsProtected and Worksheet.Protection.IsProtectedWithPassword when the password is blank or null. | Provide best‑practice guidance for handling worksheet protection in Aspose.Cells when the password may be empty, including verification and persistence across saves.

using System;
using Aspose.Cells;

// C# example that protects the first worksheet of a new workbook using an empty password, checks Worksheet.IsProtected and Worksheet.Protection.IsProtectedWithPassword flags, validates the empty password with VerifyPassword, saves the file, reloads it, and confirms that the protection state remains unchanged.
class ProtectWorksheetEmptyPasswordDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Protect the worksheet with an empty (blank) password
        // According to the API, a null or blank password means the sheet can be unprotected without a password
        worksheet.Protect(ProtectionType.All, "", null);

        // Observe protection properties after protecting
        Console.WriteLine("Worksheet.IsProtected: " + worksheet.IsProtected);
        Console.WriteLine("Worksheet.Protection.IsProtectedWithPassword: " + worksheet.Protection.IsProtectedWithPassword);

        // Verify the empty password using VerifyPassword method
        bool isPasswordCorrect = worksheet.Protection.VerifyPassword("");
        Console.WriteLine("Verify empty password (should be true): " + isPasswordCorrect);

        // Save the workbook to a file
        string fileName = "EmptyPasswordProtected.xlsx";
        workbook.Save(fileName);

        // Load the saved workbook to confirm that protection state persists
        Workbook loadedWorkbook = new Workbook(fileName);
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

        // Check protection properties on the loaded worksheet
        Console.WriteLine("Loaded Worksheet.IsProtected: " + loadedWorksheet.IsProtected);
        Console.WriteLine("Loaded Worksheet.Protection.IsProtectedWithPassword: " + loadedWorksheet.Protection.IsProtectedWithPassword);
        bool loadedVerify = loadedWorksheet.Protection.VerifyPassword("");
        Console.WriteLine("Loaded verify empty password (should be true): " + loadedVerify);
    }
}
