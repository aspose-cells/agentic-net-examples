// Title: Apply password protection that expires after a set interval to an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect worksheet with password that expires after minutes | set expiration for Excel sheet protection using Aspose.Cells .NET | remove worksheet protection automatically after a time interval in C# | store worksheet protection expiry timestamp in custom document property Aspose.Cells | test if Excel worksheet protection has expired with Aspose.Cells API
// Tags: worksheet password protection expiration Aspose.Cells | custom document property for protection expiry .NET | programmatic unprotect sheet after timeout C# | temporary Excel sheet protection using Aspose.Cells | time‑based sheet protection removal .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, applies password protection to a worksheet, records an expiration DateTime in a custom document property, saves the file, reloads it, checks whether the current time exceeds the stored expiration, and automatically removes the protection if it has expired before saving the updated workbook.
class WorksheetProtectionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "ProtectedSheet";

            // Add sample data
            ws.Cells["A1"].PutValue("Sensitive Data");
            ws.Cells["A2"].PutValue(12345);

            // Apply protection with a password (oldPassword not required, pass null)
            string password = "Secret123";
            ws.Protect(ProtectionType.All, password, null);

            // Define expiration time (e.g., 10 seconds from now)
            DateTime expiration = DateTime.Now.AddSeconds(10);

            // Store expiration time in a custom document property
            wb.CustomDocumentProperties.Add("SheetProtectionExpiration", expiration);

            // Save the workbook
            string originalPath = "ProtectedWorkbook.xlsx";
            wb.Save(originalPath);

            // ------------------- Test expiration behavior -------------------
            // Ensure the file exists before loading
            if (!File.Exists(originalPath))
                throw new FileNotFoundException($"File not found: {originalPath}");

            // Load the workbook
            Workbook loadedWb = new Workbook(originalPath);
            Worksheet loadedWs = loadedWb.Worksheets["ProtectedSheet"];

            // Retrieve the stored expiration time
            var prop = loadedWb.CustomDocumentProperties["SheetProtectionExpiration"];
            if (prop == null)
                throw new InvalidOperationException("Expiration property not found.");

            DateTime exp = (DateTime)prop.Value;

            // Check if the protection has expired
            if (DateTime.Now > exp)
            {
                // Expired: remove protection
                loadedWs.Unprotect(password);
                Console.WriteLine("Protection expired; sheet is now unprotected.");
            }
            else
            {
                Console.WriteLine($"Protection still active. Expires at: {exp}");
            }

            // Save any changes after the check
            string updatedPath = "ProtectedWorkbook_Updated.xlsx";
            loadedWb.Save(updatedPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
