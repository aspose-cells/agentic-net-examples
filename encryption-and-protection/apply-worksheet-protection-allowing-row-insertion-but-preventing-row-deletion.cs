// Title: C# – Protect an Aspose.Cells worksheet to allow row insertion but block row deletion
// Description: Demonstrates how to create a workbook, enable row insertion while disabling row deletion on a protected worksheet, optionally set a password, apply full protection with ProtectionType.All, and save the file as WorksheetProtection_InsertionOnly.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet protection | AllowInsertingRow C# | AllowDeletingRow example | protect sheet with password Aspose.Cells | row insertion only protection | prevent row deletion Aspose.Cells | ProtectionType.All usage | Aspose.Cells .NET sample
// Common Searches: Aspose.Cells protect worksheet allow insert rows | How to block row deletion in Aspose.Cells C# | Enable row insertion on a protected sheet Aspose.Cells | Set password for worksheet protection Aspose.Cells .NET | C# example for worksheet protection with row insertion only
// Developer Intent: The developer needs a protected worksheet that permits adding rows but prevents removing existing rows.
// Use Cases: Data‑entry template where users can add new records without deleting the original rows, preserving audit trails. | Report workbook that lets analysts expand sections with extra rows while locking the core layout against accidental deletions.
// AI Prompts: Generate C# code with Aspose.Cells that protects a worksheet, enables row insertion, disables row deletion, and applies an optional password. | Explain the interaction between Protection.AllowInsertingRow, Protection.AllowDeletingRow, and worksheet.Protect(ProtectionType.All) in Aspose.Cells. | Show how to save a workbook with only row insertion allowed, then reopen it to confirm the protection settings programmatically.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, enable row insertion while disabling row deletion on a protected worksheet, optionally set a password, apply full protection with ProtectionType.All, and save the file as WorksheetProtection_InsertionOnly.xlsx using Aspose.Cells for .NET.
    public class WorksheetProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the protection settings of the worksheet
                Protection protection = worksheet.Protection;

                // Allow inserting rows while the sheet is protected
                protection.AllowInsertingRow = true;

                // Disallow deleting rows while the sheet is protected
                protection.AllowDeletingRow = false;

                // Set a password for the protection (optional)
                protection.Password = "mySecretPassword";

                // Apply protection to the worksheet (protect all aspects)
                worksheet.Protect(ProtectionType.All);

                // Save the workbook to a file
                workbook.Save("WorksheetProtection_InsertionOnly.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetProtectionDemo.Run();
        }
    }
}
