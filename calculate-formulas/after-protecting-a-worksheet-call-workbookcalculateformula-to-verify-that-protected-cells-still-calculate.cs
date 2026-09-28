// Title: Calculate formulas on a password‑protected worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, assigns formulas to cells, protects the worksheet with a password, and then calls Workbook.CalculateFormula to compute the results. | Show how to verify a worksheet's protection status and still evaluate its formulas using Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells C# calculate formulas after protecting a worksheet with a password | How to run Workbook.CalculateFormula on a protected sheet in .NET | C# example for protecting Excel sheet and still evaluating cell formulas using Aspose.Cells | Retrieve calculated values from a password‑protected worksheet with Aspose.Cells
// Tags: protect worksheet password Aspose.Cells C# | Workbook.CalculateFormula on protected sheet | evaluate Excel formulas after sheet protection | Aspose.Cells worksheet protection example | C# calculate formulas in protected workbook

using System;
using Aspose.Cells;

// The sample creates a new workbook, sets values and formulas in cells, protects the first worksheet with a password, confirms the protection state, invokes Workbook.CalculateFormula to evaluate the formulas despite protection, prints the calculated results, and saves the file as ProtectedCalc.xlsx.
class ProtectAndCalculateDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set initial values and formulas
        cells["A1"].PutValue(5);
        cells["B1"].Formula = "=A1*2";
        cells["C1"].Formula = "=B1+10";

        // Protect the worksheet with a password
        sheet.Protect(ProtectionType.All, "pwd123", null);

        // Verify that the worksheet is protected
        Console.WriteLine("Worksheet protected: " + sheet.IsProtected);

        // Calculate formulas even though the sheet is protected
        workbook.CalculateFormula();

        // Output the calculated results
        Console.WriteLine("A1 value: " + cells["A1"].IntValue);
        Console.WriteLine("B1 calculated: " + cells["B1"].IntValue);
        Console.WriteLine("C1 calculated: " + cells["C1"].IntValue);

        // Save the workbook (optional)
        workbook.Save("ProtectedCalc.xlsx");
    }
}
