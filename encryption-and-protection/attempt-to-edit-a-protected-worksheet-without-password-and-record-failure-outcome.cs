// Title: Aspose.Cells .NET: Attempt to Edit a Password‑Protected Worksheet and Capture the Exception
// Description: This C# example creates a workbook, protects the first worksheet with a password, then tries to write to cell A1 without providing the password. The protected sheet blocks the edit, causing Aspose.Cells to throw an exception, which is caught and logged before the file is saved.
// Keywords: Aspose.Cells .NET | protected worksheet | worksheet protection | edit locked cell | exception handling | PutValue error | C# Excel security | catch Aspose.Cells exception | password‑protected sheet | Excel protection demo
// Common Searches: Aspose.Cells edit protected worksheet without password | C# catch exception when writing to locked cell | how to test worksheet protection in Aspose.Cells | Aspose.Cells PutValue throws error on protected sheet | example of protecting a worksheet and attempting unauthorized edit
// Developer Intent: Show that attempting to modify a locked cell on a password‑protected worksheet raises an exception and demonstrate how to record the failure.
// Use Cases: Automated verification that worksheet protection blocks unauthorized edits. | Logging attempts to modify protected Excel files in server‑side processing. | Creating unit tests for security compliance of generated workbooks.
// AI Prompts: Write code that checks Worksheet.IsProtected before calling PutValue and skips the operation if protection is active. | Provide a sample that unprotects a worksheet with a known password, updates a cell, and reapplies protection using Aspose.Cells for .NET. | Explain which Aspose.Cells exception type is thrown when PutValue is used on a locked cell in a protected worksheet.

using Aspose.Cells;
using System;

// This C# example creates a workbook, protects the first worksheet with a password, then tries to write to cell A1 without providing the password. The protected sheet blocks the edit, causing Aspose.Cells to throw an exception, which is caught and logged before the file is saved.
class EditProtectedWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet with a password
        sheet.Protect(ProtectionType.All, "secret", null);
        Console.WriteLine("Worksheet protected: " + sheet.IsProtected);

        // Attempt to edit a cell without providing the password
        try
        {
            // This should fail because the worksheet is protected and the cell is locked by default
            sheet.Cells["A1"].PutValue("Attempted edit");
            Console.WriteLine("Edit succeeded unexpectedly.");
        }
        catch (Exception ex)
        {
            // Record the failure outcome
            Console.WriteLine("Failed to edit protected worksheet without password: " + ex.Message);
        }

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("ProtectedEditAttempt.xlsx");
    }
}
