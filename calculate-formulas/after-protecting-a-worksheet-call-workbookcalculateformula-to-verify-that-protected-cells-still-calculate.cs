// Title: C# – Protect a Worksheet and Recalculate Formulas with Aspose.Cells
// Description: Demonstrates how to protect an Excel worksheet with a password, then call Workbook.CalculateFormula to evaluate all formulas on the protected sheet and save the result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# protect worksheet | Workbook.CalculateFormula | formula evaluation on protected sheet | password‑protected Excel file | read‑only report generation | Excel automation Aspose.Cells | recalculate formulas after protection | C# Excel workbook example
// Common Searches: Aspose.Cells calculate formulas after protecting worksheet | C# protect Excel sheet and still evaluate formulas | Workbook.CalculateFormula on a locked worksheet | How to recalculate formulas on a password‑protected sheet using Aspose.Cells
// Developer Intent: Ensure that formulas continue to compute after a worksheet is secured with a password.
// Use Cases: Create a read‑only financial model that updates automatically when source data changes. | Distribute a password‑protected report that always shows the latest calculated values. | Generate Excel files for external partners where editing is blocked but formulas recalculate on open.
// AI Prompts: Show C# code that protects an Excel worksheet with a password and then runs Workbook.CalculateFormula using Aspose.Cells. | Provide an Aspose.Cells example that verifies formula results after worksheet protection. | Explain the behavior of Workbook.CalculateFormula on a protected sheet and any relevant options.

using System;
using Aspose.Cells;

// Demonstrates how to protect an Excel worksheet with a password, then call Workbook.CalculateFormula to evaluate all formulas on the protected sheet and save the result using Aspose.Cells for .NET.
class ProtectAndCalculateDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate cells with values and formulas
        cells["A1"].PutValue(5);                 // Base value
        cells["B1"].Formula = "=A1*2";           // Should evaluate to 10
        cells["C1"].Formula = "=B1+10";          // Should evaluate to 20

        // Protect the worksheet with a password
        sheet.Protect(ProtectionType.All, "pwd123", null);

        // Confirm that the worksheet is protected
        Console.WriteLine("Worksheet protected: " + sheet.IsProtected);

        // Calculate all formulas after protection
        workbook.CalculateFormula();

        // Display the results to verify calculation works on protected cells
        Console.WriteLine("A1 value: " + cells["A1"].IntValue);
        Console.WriteLine("B1 formula result: " + cells["B1"].IntValue);
        Console.WriteLine("C1 formula result: " + cells["C1"].IntValue);

        // Save the workbook (optional)
        workbook.Save("ProtectedCalc.xlsx");
    }
}
