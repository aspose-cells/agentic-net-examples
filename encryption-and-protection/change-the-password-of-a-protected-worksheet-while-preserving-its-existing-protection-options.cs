// Title: Change a protected worksheet password while retaining protection settings – Aspose.Cells for .NET
// Description: Demonstrates how to replace the password of a protected worksheet using Aspose.Cells for .NET without altering any existing protection flags. The example creates a workbook, sets custom protection options, applies an initial password, swaps it for a new one via the Protect overload that accepts old and new passwords, verifies the change, and saves the file.
// Keywords: Aspose.Cells change worksheet password | preserve worksheet protection options | Aspose.Cells Protect overload old new password | update Excel sheet password .NET | worksheet unprotect with new password
// Common Searches: how to change worksheet password in Aspose.Cells without losing protection settings | Aspose.Cells replace worksheet password while keeping AllowDeletingRow flag | C# change Excel worksheet password programmatically Aspose.Cells | protect worksheet with new password overload Aspose.Cells
// Developer Intent: Replace the current worksheet password with a new one while keeping all existing protection settings unchanged.
// Use Cases: Rotate worksheet passwords for compliance without resetting locked cells or formatting rules. | Automate password updates across multiple workbooks before distribution to different user groups. | Update a workbook’s password prior to publishing while preserving custom protection flags such as AllowDeletingColumn.
// AI Prompts: Write C# code that changes a protected worksheet's password using Aspose.Cells and keeps all protection options intact. | Explain how the Protect method overload that takes old and new passwords retains worksheet protection settings. | Provide error‑handling best practices for confirming that a worksheet can be unprotected with the new password after a password change.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to replace the password of a protected worksheet using Aspose.Cells for .NET without altering any existing protection flags. The example creates a workbook, sets custom protection options, applies an initial password, swaps it for a new one via the Protect overload that accepts old and new passwords, verifies the change, and saves the file.
    public class ChangeWorksheetPassword
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set some protection options
                Protection protection = sheet.Protection;
                protection.AllowEditingObject = false;
                protection.AllowEditingScenario = false;
                protection.AllowDeletingColumn = true;
                protection.AllowDeletingRow = true;

                // Initial protection with a password
                string oldPassword = "oldPass123";
                sheet.Protect(ProtectionType.All, oldPassword, null);

                // Verify that the worksheet is protected
                Console.WriteLine("IsProtected: " + sheet.IsProtected);
                Console.WriteLine("IsProtectedWithPassword: " + sheet.Protection.IsProtectedWithPassword);

                // Change the password while preserving all protection options
                string newPassword = "newPass456";
                // The overload takes the old password and the new password, keeping existing options intact
                sheet.Protect(ProtectionType.All, newPassword, oldPassword);

                // Verify that the new password works
                bool canUnprotect = false;
                try
                {
                    sheet.Unprotect(newPassword);
                    canUnprotect = true;
                }
                catch (Exception)
                {
                    // Unprotect failed
                }

                Console.WriteLine("Can unprotect with new password: " + canUnprotect);

                // Save the workbook
                workbook.Save("ChangedPasswordWorksheet.xlsx");
                Console.WriteLine("Workbook saved as ChangedPasswordWorksheet.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ChangeWorksheetPassword.Run();
        }
    }
}
