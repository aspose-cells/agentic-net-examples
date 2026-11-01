// Title: How to toggle Excel function language between English and French using FormulaLocal and Workbook CultureInfo in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set workbook CultureInfo to French for localized formulas | use FormulaLocal to write French SUM (SOMME) function in a .NET workbook | switch workbook culture back to invariant to use English formulas with FormulaLocal | example of toggling between Formula and FormulaLocal properties in Aspose.Cells | C# code to change Excel function language dynamically with Aspose.Cells
// Tags: FormulaLocal with workbook CultureInfo | localized Excel functions Aspose.Cells .NET | toggle workbook culture for formulas C# | invariant culture usage FormulaLocal | French SOMME function Aspose.Cells example

using System;
using System.Globalization;
using Aspose.Cells;

// The sample creates a workbook, adds numeric data, assigns an English SUM formula via the Formula property, changes the workbook's CultureInfo to French and sets a localized SOMME formula using FormulaLocal, then resets the culture to InvariantCulture and applies an English SUM formula again with FormulaLocal, finally saving the file as FormulaLocalDemo.xlsx.
class FormulaLocalExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Access the first worksheet
        Worksheet ws = wb.Worksheets[0];

        // Populate some sample data for the SUM function
        ws.Cells["B1"].PutValue(10);
        ws.Cells["B2"].PutValue(20);
        ws.Cells["B3"].PutValue(30);
        ws.Cells["B4"].PutValue(40);
        ws.Cells["B5"].PutValue(50);

        // -----------------------------------------------------------------
        // 1. English function name using the Formula property (always English)
        // -----------------------------------------------------------------
        ws.Cells["A1"].Formula = "=SUM(B1:B5)";

        // -----------------------------------------------------------------
        // 2. Localized function name using FormulaLocal
        //    Set workbook culture to French (fr-FR). In French, SUM is "SOMME".
        // -----------------------------------------------------------------
        wb.Settings.CultureInfo = new CultureInfo("fr-FR");
        ws.Cells["A2"].FormulaLocal = "=SOMME(B1:B5)";

        // -----------------------------------------------------------------
        // 3. Switch back to English (InvariantCulture) and use FormulaLocal
        //    FormulaLocal now accepts English function names again.
        // -----------------------------------------------------------------
        wb.Settings.CultureInfo = CultureInfo.InvariantCulture;
        ws.Cells["A3"].FormulaLocal = "=SUM(B1:B5)";

        // Save the workbook to a file
        wb.Save("FormulaLocalDemo.xlsx");
    }
}
