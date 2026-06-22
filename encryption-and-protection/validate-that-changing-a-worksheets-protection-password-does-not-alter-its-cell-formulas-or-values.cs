using System;
using Aspose.Cells;

class WorksheetProtectionPasswordChangeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells with values and a formula
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Store original cell contents for later comparison
        object originalValueA1 = sheet.Cells["A1"].Value;
        object originalValueA2 = sheet.Cells["A2"].Value;
        string originalFormulaA3 = sheet.Cells["A3"].Formula;
        workbook.CalculateFormula(); // Ensure formula is evaluated
        object originalValueA3 = sheet.Cells["A3"].Value;

        // Protect the worksheet with an initial password
        sheet.Protect(ProtectionType.All, "oldpwd", null);

        // Change the protection password from "oldpwd" to "newpwd"
        sheet.Protect(ProtectionType.All, "newpwd", "oldpwd");

        // Verify that the old password no longer works and the new one does
        bool oldPasswordValid = sheet.Protection.VerifyPassword("oldpwd"); // expected: false
        bool newPasswordValid = sheet.Protection.VerifyPassword("newpwd"); // expected: true

        Console.WriteLine($"Old password valid: {oldPasswordValid}");
        Console.WriteLine($"New password valid: {newPasswordValid}");

        // Recalculate formulas to ensure values are up‑to‑date
        workbook.CalculateFormula();

        // Validate that cell values and formulas are unchanged after password change
        bool a1Unchanged = object.Equals(originalValueA1, sheet.Cells["A1"].Value);
        bool a2Unchanged = object.Equals(originalValueA2, sheet.Cells["A2"].Value);
        bool a3FormulaUnchanged = originalFormulaA3 == sheet.Cells["A3"].Formula;
        bool a3ValueUnchanged = object.Equals(originalValueA3, sheet.Cells["A3"].Value);

        Console.WriteLine($"A1 unchanged: {a1Unchanged}");
        Console.WriteLine($"A2 unchanged: {a2Unchanged}");
        Console.WriteLine($"A3 formula unchanged: {a3FormulaUnchanged}");
        Console.WriteLine($"A3 value unchanged: {a3ValueUnchanged}");

        // Save the workbook (optional)
        workbook.Save("PasswordChanged.xlsx");
    }
}