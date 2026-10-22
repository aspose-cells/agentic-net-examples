// Title: How to add password write protection to an Excel worksheet and then remove it with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells protect worksheet with password and save workbook | remove password protection from Excel sheet using Aspose.Cells C# | unprotect worksheet after loading a password‑protected workbook with Aspose.Cells | apply and later clear write protection on an Excel worksheet programmatically | Aspose.Cells Protect and Unprotect methods example in .NET
// Tags: Aspose.Cells worksheet protect method C# | Aspose.Cells worksheet unprotect method C# | password protected Excel workbook Aspose.Cells | remove write protection from Excel sheet Aspose.Cells | save and reload protected workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new workbook, applies write protection to the first worksheet using ws.Protect with a password, saves the protected file, reloads it, removes the protection with ws.Unprotect using the same password, and saves the resulting unprotected workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet and rename it
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "ProtectedSheet";

            // Apply write protection with a password (oldPassword not required)
            string password = "MySecret123";
            ws.Protect(ProtectionType.All, password, null);

            // Save the protected workbook
            string protectedPath = "Protected.xlsx";
            wb.Save(protectedPath);

            // Verify the file exists before loading
            if (!File.Exists(protectedPath))
                throw new FileNotFoundException($"File not found: {protectedPath}");

            // Load the workbook to demonstrate removal of protection
            Workbook wb2 = new Workbook(protectedPath);

            // Access the protected worksheet
            Worksheet ws2 = wb2.Worksheets[0];

            // Remove protection using the same password
            ws2.Unprotect(password);

            // Save the unprotected workbook
            string unprotectedPath = "Unprotected.xlsx";
            wb2.Save(unprotectedPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
