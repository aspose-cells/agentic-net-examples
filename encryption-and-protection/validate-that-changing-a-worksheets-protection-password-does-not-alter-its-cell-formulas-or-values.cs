// Title: Aspose.Cells for .NET – Verify that Changing a Worksheet Protection Password Preserves Formulas and Values
// Description: C# example that creates a workbook, adds numeric data and a formula, protects the sheet with an initial password, switches to a new password, and confirms that the original formula and calculated value remain unchanged before saving the file.
// Keywords: Aspose.Cells worksheet protection password change | C# preserve cell formulas after unprotect | Aspose.Cells keep calculated values intact | protect sheet with new password .NET | worksheet password rotation example | formula integrity after re‑protecting | Aspose.Cells demo password update | verify cell value stability Aspose
// Common Searches: change worksheet protection password without altering formulas Aspose.Cells | does unprotecting a sheet affect calculated values .NET | how to re‑protect an Aspose.Cells worksheet with a new password | verify formula consistency after password change Aspose | Aspose.Cells keep cell values after protecting sheet
// Developer Intent: Confirm that updating a worksheet's protection password does not modify any existing formulas or their calculated results.
// Use Cases: Automated password rotation for financial models while guaranteeing formula accuracy. | Unit testing of reporting dashboards to ensure data integrity after re‑protecting sheets. | Batch processing of multiple worksheets to apply new passwords without corrupting calculations.
// AI Prompts: Generate C# code using Aspose.Cells that changes a worksheet's protection password and asserts that all formulas and values stay unchanged. | Create an NUnit test that validates formula and value consistency after updating the worksheet protection password with Aspose.Cells. | Explain Aspose.Cells' behavior regarding formula recalculation when a worksheet is unprotected and then re‑protected with a different password.

using System;
using Aspose.Cells;

namespace WorksheetProtectionPasswordChangeDemo
{
    // C# example that creates a workbook, adds numeric data and a formula, protects the sheet with an initial password, switches to a new password, and confirms that the original formula and calculated value remain unchanged before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with values and a formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].Formula = "=A1+B1";

            // Store original value and formula for later comparison
            object originalValue = sheet.Cells["C1"].Value;
            string originalFormula = sheet.Cells["C1"].Formula;

            // Protect the worksheet with an initial password
            string oldPassword = "oldPass";
            sheet.Protect(ProtectionType.All, oldPassword, null);

            // Change the protection password:
            // 1. Unprotect using the old password
            sheet.Unprotect(oldPassword);
            // 2. Protect again with a new password
            string newPassword = "newPass";
            sheet.Protect(ProtectionType.All, newPassword, null);

            // After changing the password, read the cell's value and formula again
            object afterValue = sheet.Cells["C1"].Value;
            string afterFormula = sheet.Cells["C1"].Formula;

            // Verify that the value and formula have not changed
            bool valueUnchanged = Equals(originalValue, afterValue);
            bool formulaUnchanged = string.Equals(originalFormula, afterFormula, StringComparison.Ordinal);

            Console.WriteLine($"Value unchanged: {valueUnchanged}");
            Console.WriteLine($"Formula unchanged: {formulaUnchanged}");

            // Save the workbook (optional, just to complete lifecycle)
            workbook.Save("WorksheetProtectionPasswordChangeDemo.xlsx");
        }
    }
}
