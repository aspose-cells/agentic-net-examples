// Title: Check that updating a worksheet's protection password does not modify formulas or calculated values with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to protect a worksheet, change its password, recalculate formulas, and assert that the original formula text and result remain unchanged. | Create a .NET example that saves a workbook to a MemoryStream after changing the worksheet protection password, reloads it, and verifies that the formula and its evaluated value are preserved.
// Common Searches: Aspose.Cells C# verify formula remains same after changing worksheet password | how to change worksheet protection password without affecting calculated cells in .NET | test if unprotect and protect worksheet alters cell formulas Aspose.Cells | preserve Excel formula values when updating sheet protection password using Aspose.Cells | reload workbook from stream to confirm formula persistence after password change
// Tags: worksheet protection password change Aspose.Cells | formula integrity after unprotect protect | recalculate formulas after worksheet password update | save workbook to memory stream Aspose.Cells | load workbook from stream verify formula

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills cells, sets a formula, protects the sheet, changes the protection password by unprotecting and re‑protecting, recalculates, and confirms that both the formula text and its evaluated value stay unchanged. It then saves the workbook to a MemoryStream, reloads it, and validates that the formula integrity persists after the password update.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate cells with data
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(10);

            // Set a formula in C1
            sheet.Cells["C1"].Formula = "=A1+B1";

            // Calculate formulas to obtain values
            workbook.CalculateFormula();

            // Store original formula and evaluated value
            string originalFormula = sheet.Cells["C1"].Formula;
            object originalValue = sheet.Cells["C1"].Value;

            // Protect the worksheet with an initial password (oldPassword not required, pass null)
            sheet.Protect(ProtectionType.All, "oldpwd", null);

            // Change the protection password: unprotect then protect with a new password
            sheet.Unprotect("oldpwd");
            sheet.Protect(ProtectionType.All, "newpwd", null);

            // Recalculate to ensure formulas are still valid
            workbook.CalculateFormula();

            // Verify that the formula and its value have not changed
            bool formulaUnchanged = sheet.Cells["C1"].Formula == originalFormula;
            bool valueUnchanged = Equals(sheet.Cells["C1"].Value, originalValue);

            Console.WriteLine($"Formula unchanged after password change: {formulaUnchanged}");
            Console.WriteLine($"Value unchanged after password change: {valueUnchanged}");

            // Save the workbook to a memory stream (demonstrating the save rule)
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0;

                // Load the workbook from the memory stream (demonstrating the load rule)
                var loadedWorkbook = new Workbook(ms);
                var loadedSheet = loadedWorkbook.Worksheets[0];

                // Verify again after load to ensure persistence
                bool formulaAfterLoad = loadedSheet.Cells["C1"].Formula == originalFormula;
                bool valueAfterLoad = Equals(loadedSheet.Cells["C1"].Value, originalValue);

                Console.WriteLine($"Formula unchanged after load: {formulaAfterLoad}");
                Console.WriteLine($"Value unchanged after load: {valueAfterLoad}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
