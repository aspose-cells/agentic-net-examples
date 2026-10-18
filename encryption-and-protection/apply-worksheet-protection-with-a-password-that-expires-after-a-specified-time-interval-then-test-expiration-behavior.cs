// Title: Protect an Excel worksheet with a password that expires after a set interval using Aspose.Cells for .NET (C#)
// Common Searches: how to set a timed password for worksheet protection with Aspose.Cells in C# | Aspose.Cells example for worksheet protection that expires after seconds | test unprotecting an Excel sheet after password expiration using Aspose.Cells | C# code to apply and verify expiration of worksheet protection in Aspose.Cells
// Tags: worksheet protection expiration Aspose.Cells | timed password protection Excel C# | Aspose.Cells unprotect worksheet after timeout | ProtectionType.All usage Aspose.Cells | save protected workbook Aspose.Cells

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The sample creates a workbook, protects the first worksheet with a password using ProtectionType.All, records an expiration time based on a configurable interval, demonstrates successful unprotect before expiration, waits past the interval, shows that unprotect is denied after expiration, and finally saves the workbook.
class WorksheetProtectionWithExpiration
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Add some sample data
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["A2"].PutValue(12345);

            // Define protection password and expiration interval (in seconds)
            string password = "SecretPwd";
            int expirationSeconds = 5;

            // Apply worksheet protection with the password (oldPassword not required, pass empty string)
            sheet.Protect(ProtectionType.All, password, string.Empty);
            Console.WriteLine("Worksheet protected with password.");

            // Store the expiration time
            DateTime expirationTime = DateTime.Now.AddSeconds(expirationSeconds);
            Console.WriteLine($"Protection will expire at: {expirationTime:HH:mm:ss}");

            // ----- Test before expiration -----
            Console.WriteLine("\n--- Testing before expiration ---");
            TryUnprotect(sheet, password, expirationTime);

            // Wait until after the expiration time
            int waitTime = expirationSeconds + 2; // extra buffer
            Console.WriteLine($"\nWaiting {waitTime} seconds to exceed expiration...");
            Thread.Sleep(waitTime * 1000);

            // ----- Test after expiration -----
            Console.WriteLine("\n--- Testing after expiration ---");
            TryUnprotect(sheet, password, expirationTime);

            // Save the workbook (optional)
            string outputPath = "ProtectedSheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to attempt unprotecting the worksheet respecting expiration
    static void TryUnprotect(Worksheet sheet, string password, DateTime expiration)
    {
        try
        {
            if (DateTime.Now <= expiration)
            {
                // Within valid period: unprotect succeeds
                sheet.Unprotect(password);
                Console.WriteLine("Unprotect succeeded (within expiration).");

                // Re‑apply protection to keep the sheet protected for further tests
                sheet.Protect(ProtectionType.All, password, string.Empty);
            }
            else
            {
                // After expiration: do not allow unprotect
                Console.WriteLine("Unprotect denied: protection has expired.");
                // Optionally, automatically remove protection after expiration
                // sheet.Unprotect(password);
                // Console.WriteLine("Protection automatically removed after expiration.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unprotect attempt failed: {ex.Message}");
        }
    }
}
