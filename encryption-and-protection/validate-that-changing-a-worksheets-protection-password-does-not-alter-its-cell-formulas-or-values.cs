// Title: Check that changing a worksheet protection password preserves formulas and values – Aspose.Cells for .NET
// Description: Demonstrates how to protect a worksheet, change its password, verify the new password, and confirm that cell contents—including formulas and calculated results—remain unchanged after saving and reloading the workbook.
// Keywords: Aspose.Cells worksheet password change | preserve formulas after protection update | verify cell values after password rotation | worksheet protection .NET | Excel password change data integrity
// Common Searches: change worksheet protection password without losing data Aspose.Cells | does updating Excel sheet password affect formulas | verify data integrity after worksheet password change .NET | how to rotate worksheet protection password in Aspose.Cells
// Developer Intent: Confirm that modifying a worksheet's protection password does not alter any cell formulas or values.
// Use Cases: Automate password rotation for secured worksheets while guaranteeing calculation results stay intact. | Create a regression test that validates data integrity after changing worksheet protection passwords. | Implement a maintenance routine that updates worksheet passwords in bulk without risking formula corruption.
// AI Prompts: Generate a C# unit test using Aspose.Cells that changes a worksheet's protection password and asserts that all cell values and formulas are unchanged. | Write code to programmatically replace a worksheet protection password, verify the new password, and confirm no data loss. | Explain the mechanism Aspose.Cells uses to keep cell contents stable when a worksheet's protection password is updated.

using System;
using Aspose.Cells;

// Demonstrates how to protect a worksheet, change its password, verify the new password, and confirm that cell contents—including formulas and calculated results—remain unchanged after saving and reloading the workbook.
class WorksheetProtectionPasswordChangeDemo
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with values and a formula
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Calculate formulas so that values are materialized
        workbook.CalculateFormula();

        // Store original cell values and formula for later comparison
        object originalA1 = worksheet.Cells["A1"].Value;
        object originalA2 = worksheet.Cells["A2"].Value;
        string originalFormulaA3 = worksheet.Cells["A3"].Formula;
        object originalA3 = worksheet.Cells["A3"].Value;

        // Protect the worksheet with an initial password
        worksheet.Protect(ProtectionType.All, "oldPass", null);

        // Change the protection password from "oldPass" to "newPass"
        worksheet.Protect(ProtectionType.All, "newPass", "oldPass");

        // Verify that the new password is valid
        bool isNewPasswordCorrect = worksheet.Protection.VerifyPassword("newPass");
        Console.WriteLine("New password verification result: " + isNewPasswordCorrect);

        // Save the workbook after password change
        workbook.Save("PasswordChanged.xlsx");

        // Load the saved workbook to ensure data integrity
        Workbook loadedWorkbook = new Workbook("PasswordChanged.xlsx");
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

        // Recalculate formulas in the loaded workbook
        loadedWorkbook.CalculateFormula();

        // Compare cell values and formula with the originals
        bool dataUnchanged =
            loadedWorksheet.Cells["A1"].Value.Equals(originalA1) &&
            loadedWorksheet.Cells["A2"].Value.Equals(originalA2) &&
            loadedWorksheet.Cells["A3"].Formula == originalFormulaA3 &&
            loadedWorksheet.Cells["A3"].Value.Equals(originalA3);

        Console.WriteLine("Cell values and formulas unchanged after password change: " + dataUnchanged);
    }
}
