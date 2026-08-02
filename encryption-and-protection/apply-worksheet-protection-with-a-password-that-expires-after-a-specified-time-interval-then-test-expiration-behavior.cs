// Title: Timed Worksheet Protection with Auto‑Unprotect in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to protect an Excel worksheet with a password, store an expiration UTC timestamp in a hidden column, save the file, wait past the interval, reload the workbook, compare the current time with the stored ticks, and automatically unprotect the sheet when the password has expired.
// Keywords: Aspose.Cells | C# | worksheet protection | password expiration | auto unprotect | hidden cell | expiration timestamp | UTC ticks | timed Excel security | Excel sheet lock
// Common Searches: Aspose.Cells protect worksheet with expiration time | C# auto‑unprotect Excel sheet after timeout | store expiration date in hidden column Aspose.Cells | check worksheet protection expiry in .NET | timed password for Excel worksheet using Aspose
// Developer Intent: The developer needs to apply a password to a worksheet that becomes invalid after a set period and to programmatically remove the protection once the expiration time is reached.
// Use Cases: Temporarily lock confidential data while allowing limited‑time access. | Hide the expiry information from end users by storing it in a concealed column. | Automatically restore editability when a workbook is opened after the defined timeout.
// AI Prompts: Generate C# code that protects an Aspose.Cells worksheet with a password, writes an expiration DateTime (as UTC ticks) to a hidden cell, and saves the workbook. | Create a method that loads a protected workbook, reads the expiration ticks from the hidden cell, determines if the current UTC time exceeds the stored value, and calls Unprotect if the password has expired. | Provide error‑handling examples for missing or malformed expiration data when validating timed worksheet protection.

using System;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtectionWithExpiration
{
    // Demonstrates how to protect an Excel worksheet with a password, store an expiration UTC timestamp in a hidden column, save the file, wait past the interval, reload the workbook, compare the current time with the stored ticks, and automatically unprotect the sheet when the password has expired.
    class Program
    {
        static void Main()
        {
            // ---------- Create and protect worksheet ----------
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put some sample data
            sheet.Cells["B2"].PutValue("Sensitive Data");

            // Define protection password and expiration interval (e.g., 30 seconds)
            string password = "SecretPwd123";
            TimeSpan expirationInterval = TimeSpan.FromSeconds(30);

            // Calculate expiration DateTime and store its ticks in a hidden cell (e.g., Z1)
            DateTime expirationTime = DateTime.UtcNow.Add(expirationInterval);
            sheet.Cells["Z1"].PutValue(expirationTime.Ticks);
            // Optionally hide the column to keep it out of view
            sheet.Cells.HideColumn(25); // Column Z is index 25 (0‑based)

            // Protect the worksheet with the password (all protection types)
            sheet.Protect(ProtectionType.All, password, null);

            // Save the workbook
            string filePath = "ProtectedWithExpiration.xlsx";
            workbook.Save(filePath);
            workbook.Dispose();

            // ---------- Simulate waiting period ----------
            // Wait longer than the expiration interval to trigger expiration
            Console.WriteLine("Waiting for expiration...");
            Thread.Sleep(TimeSpan.FromSeconds(35));

            // ---------- Load workbook and test expiration ----------
            Workbook loadedWb = new Workbook(filePath);
            Worksheet loadedSheet = loadedWb.Worksheets[0];

            // Retrieve stored expiration ticks and convert back to DateTime
            long storedTicks = Convert.ToInt64(loadedSheet.Cells["Z1"].Value);
            DateTime storedExpiration = new DateTime(storedTicks, DateTimeKind.Utc);
            bool isExpired = DateTime.UtcNow > storedExpiration;

            Console.WriteLine($"Current UTC time: {DateTime.UtcNow}");
            Console.WriteLine($"Stored expiration UTC time: {storedExpiration}");
            Console.WriteLine($"Is protection expired? {isExpired}");

            if (isExpired)
            {
                // If expired, unprotect automatically using the known password
                loadedSheet.Unprotect(password);
                Console.WriteLine("Worksheet has been unprotected due to expiration.");
            }
            else
            {
                // If not expired, verify the password still works
                bool passwordCorrect = loadedSheet.Protection.VerifyPassword(password);
                Console.WriteLine($"Password verification (still protected): {passwordCorrect}");
            }

            // Save the workbook after handling expiration
            string resultPath = "ProtectedWithExpiration_Result.xlsx";
            loadedWb.Save(resultPath);
            loadedWb.Dispose();

            Console.WriteLine("Process completed.");
        }
    }
}
