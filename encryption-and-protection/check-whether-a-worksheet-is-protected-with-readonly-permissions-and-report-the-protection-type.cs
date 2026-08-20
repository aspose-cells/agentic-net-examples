// Title: C# – Detect Read‑Only Worksheet Protection and Retrieve Its Type with Aspose.Cells
// Description: This example creates a workbook, applies Contents (read‑only) protection with a password, then uses the Worksheet.IsProtected flag, the Protection.Type enum, and the Protection.IsProtectedWithPassword property to report whether the sheet is protected, which protection mode is active, and if a password is required.
// Keywords: Aspose.Cells worksheet protection check | read‑only sheet detection C# | ProtectionType.Contents example | Worksheet.IsProtected property | password‑protected worksheet Aspose
// Common Searches: how to check if an Excel sheet is read‑only protected using Aspose.Cells .NET | retrieve protection type of a worksheet in Aspose.Cells C# | determine if worksheet protection uses a password with Aspose.Cells
// Developer Intent: Identify whether a worksheet is protected with read‑only (Contents) permission, discover the exact protection type applied, and know if a password is set.
// Use Cases: Validate protection status before modifying sheets in an automated reporting pipeline. | Log protection details for each worksheet when processing user‑uploaded Excel files. | Conditionally remove protection only when it is read‑only and password‑protected.
// AI Prompts: Generate C# code with Aspose.Cells that checks if a worksheet is read‑only protected and prints the protection type and password flag. | Explain how to programmatically differentiate between Contents, Objects, and Scenarios protection types in Aspose.Cells.

using System;
using Aspose.Cells;

namespace WorksheetProtectionCheck
{
    // This example creates a workbook, applies Contents (read‑only) protection with a password, then uses the Worksheet.IsProtected flag, the Protection.Type enum, and the Protection.IsProtectedWithPassword property to report whether the sheet is protected, which protection mode is active, and if a password is required.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Protect the worksheet with read‑only (Contents) protection and a password
            worksheet.Protect(ProtectionType.Contents, "pwd123", null);

            // Check if the worksheet is protected
            bool isProtected = worksheet.IsProtected;
            Console.WriteLine($"Worksheet is protected: {isProtected}");

            // Report the protection type used (Contents corresponds to read‑only)
            Console.WriteLine($"Protection type applied: {ProtectionType.Contents}");

            // Indicate whether the protection is password‑based
            Console.WriteLine($"Protected with password: {worksheet.Protection.IsProtectedWithPassword}");
        }
    }
}
