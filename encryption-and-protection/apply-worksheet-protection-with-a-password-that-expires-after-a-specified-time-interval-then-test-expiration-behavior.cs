// Title: How to protect an Aspose.Cells worksheet with a password that expires after a set time and verify the expiration in C#
// AI Prompts: Use Aspose.Cells for .NET to protect a worksheet with a password, schedule the password to become invalid after a specified TimeSpan, and programmatically confirm that unprotect succeeds before the timeout. | Adjust the expiration interval, re‑protect the sheet with a new random password after the timeout, and verify that the original password no longer works.
// Common Searches: Aspose.Cells C# protect worksheet with temporary password that expires after minutes | how to set expiration time for Excel sheet protection using Aspose.Cells | C# example to test worksheet unprotect before and after password timeout with Aspose.Cells | simulate password expiry on an Excel worksheet in .NET | Aspose.Cells protect sheet and automatically change password after a time interval
// Tags: protect worksheet with password Aspose.Cells | time‑based worksheet protection .NET | worksheet password expiration C# | unprotect Excel sheet after timeout Aspose.Cells | simulate password expiry Aspose.Cells

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// Demonstrates using Aspose.Cells for .NET to protect a worksheet with a password, define a 5‑second expiration, save the workbook, then test that the sheet can be unprotected before the timeout and fails after the password is replaced once the interval has passed.
class WorksheetProtectionWithExpiration
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SensitiveData";

            // Put some sample data
            sheet.Cells["A1"].PutValue("Confidential Information");

            // Define a password and protect the worksheet
            string password = "Secret123";
            // The third parameter is the old password (empty because the sheet is not yet protected)
            sheet.Protect(ProtectionType.All, password, string.Empty);

            // Set an expiration interval (e.g., 5 seconds)
            TimeSpan expirationInterval = TimeSpan.FromSeconds(5);
            DateTime expirationTime = DateTime.Now.Add(expirationInterval);

            // Save the workbook (optional, just to have a file)
            string filePath = "ProtectedSheet.xlsx";
            try
            {
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }

            // ---------- Test before expiration ----------
            // Attempt to unprotect using the correct password
            bool unprotectedBefore = false;
            try
            {
                sheet.Unprotect(password);
                unprotectedBefore = true; // succeeded
            }
            catch
            {
                unprotectedBefore = false;
            }
            Console.WriteLine($"Unprotected before expiration: {unprotectedBefore}");

            // Re‑protect the sheet for the next test
            sheet.Protect(ProtectionType.All, password, string.Empty);

            // ---------- Wait until after expiration ----------
            Thread.Sleep((int)expirationInterval.TotalMilliseconds + 1000); // wait a bit longer than the interval

            // Simulate expiration: change the password to a new random one
            if (DateTime.Now > expirationTime)
            {
                // Unprotect with the old password (if still valid) and protect with a new one
                try { sheet.Unprotect(password); } catch { }
                string newPassword = Guid.NewGuid().ToString("N");
                sheet.Protect(ProtectionType.All, newPassword, string.Empty);
            }

            // Attempt to unprotect using the original (now expired) password
            bool unprotectedAfter = false;
            try
            {
                sheet.Unprotect(password);
                unprotectedAfter = true; // should not happen
            }
            catch
            {
                unprotectedAfter = false; // expected failure
            }
            Console.WriteLine($"Unprotected after expiration (should be false): {unprotectedAfter}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
