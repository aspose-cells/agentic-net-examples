// Title: Validate that changing a worksheet protection password does not alter cell formulas or values with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# verify that changing worksheet protection password does not modify formulas | Check cell values remain unchanged after updating worksheet password with Aspose.Cells | How to test worksheet protection password update without affecting calculated results in .NET | Validate formula integrity after protecting a sheet with a new password using Aspose.Cells
// Tags: worksheet protection password update validation | preserve cell formulas after sheet protect Aspose.Cells | ensure cell values unchanged on password change C# | Aspose.Cells formula consistency after password modification | C# worksheet protect without data loss

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The example creates a workbook, fills cells with values and formulas, calculates them, stores the original formulas and values, protects the worksheet with an initial password, changes the protection password, recalculates, and then iterates through all cells to compare current formulas and values with the stored originals, reporting whether any changes occurred.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill cells with values and formulas
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";          // 30
            sheet.Cells["B1"].Formula = "=SUM(A1:A2)";    // 30
            sheet.Cells["B2"].Formula = "=A3*2";          // 60

            // Calculate all formulas
            workbook.CalculateFormula();

            // Store original formulas and values for later comparison
            var originalFormulas = new Dictionary<string, string>();
            var originalValues = new Dictionary<string, object>();

            foreach (Cell cell in sheet.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                    originalFormulas[cell.Name] = cell.Formula;

                originalValues[cell.Name] = cell.Value;
            }

            // Protect the worksheet with an initial password.
            // The Protect method requires the old password parameter; use empty string for initial protection.
            sheet.Protect(ProtectionType.All, "oldpwd", string.Empty);

            // Change the protection password
            sheet.Protection.Password = "newpwd";

            // Re‑calculate to ensure no side‑effects (optional)
            workbook.CalculateFormula();

            // Verify that formulas and values have not changed
            bool validationPassed = true;

            foreach (Cell cell in sheet.Cells)
            {
                string name = cell.Name;

                // Check formula consistency
                string originalFormula = originalFormulas.ContainsKey(name) ? originalFormulas[name] : string.Empty;
                if (cell.Formula != originalFormula)
                {
                    validationPassed = false;
                    Console.WriteLine($"Formula changed in {name}: before='{originalFormula}', after='{cell.Formula}'");
                }

                // Check value consistency
                object originalValue = originalValues[name];
                if (!object.Equals(cell.Value, originalValue))
                {
                    validationPassed = false;
                    Console.WriteLine($"Value changed in {name}: before='{originalValue}', after='{cell.Value}'");
                }
            }

            Console.WriteLine(validationPassed
                ? "Validation passed: worksheet protection password change did not alter formulas or values."
                : "Validation failed: some formulas or values were altered.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
