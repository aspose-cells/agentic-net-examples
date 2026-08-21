// Title: Set German Excel formulas with Cell.FormulaLocal in Aspose.Cells for .NET
// Description: Shows how to configure Workbook.Settings.Region for Germany, place sample values, assign an English SUM formula, replace it with the German SUMME function via Cell.FormulaLocal, confirm the automatic conversion to the English syntax, calculate the workbook, and save the file.
// Keywords: Aspose.Cells | FormulaLocal | German Excel formula | SUMME | .NET | C# | Workbook.Settings.Region | localized functions | Excel localization
// Common Searches: Aspose.Cells FormulaLocal German example | How to use SUMME function with Aspose.Cells C# | Set workbook region to Germany Aspose.Cells | Convert English formula to local language in Aspose.Cells | Calculate workbook after applying localized formula
// Developer Intent: Write a formula in German using Cell.FormulaLocal, let Aspose.Cells translate it to the internal English representation, and ensure correct calculation and saving of the workbook.
// Use Cases: Set Workbook.Settings.Region = CountryCode.Germany to enable German function names. | Apply Cell.FormulaLocal = "=SUMME(B1:C1)" to store a German formula. | Read Cell.Formula to see the automatically translated English formula (=SUM(B1:C1)). | Call Workbook.CalculateFormula() to evaluate the localized formula. | Save the workbook while preserving the German formula syntax.
// AI Prompts: Generate C# code that uses Cell.FormulaLocal to set a French SOMME formula in Aspose.Cells and displays the translated English formula. | Explain how Workbook.Settings.Region influences formula translation with Cell.FormulaLocal in Aspose.Cells, including calculation behavior and file output.

using System;
using Aspose.Cells;

namespace AsposeCellsLocalizedFormulaDemo
{
    // Shows how to configure Workbook.Settings.Region for Germany, place sample values, assign an English SUM formula, replace it with the German SUMME function via Cell.FormulaLocal, confirm the automatic conversion to the English syntax, calculate the workbook, and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Set the workbook region to Germany to demonstrate German localized functions
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put some sample values that the formulas will use
            cells["B1"].PutValue(5);
            cells["C1"].PutValue(10);

            // Set a formula using the standard (English) function name
            Cell cellA1 = cells["A1"];
            cellA1.Formula = "=SUM(B1:C1)";

            // Display the standard formula
            Console.WriteLine("Standard Formula: " + cellA1.Formula);

            // Set the same formula using the localized (German) function name via FormulaLocal
            cellA1.FormulaLocal = "=SUMME(B1:C1)";

            // Display the localized formula
            Console.WriteLine("Localized Formula (FormulaLocal): " + cellA1.FormulaLocal);

            // Verify that the standard Formula property reflects the English name after setting FormulaLocal
            Console.WriteLine("Standard Formula after setting FormulaLocal: " + cellA1.Formula);

            // Calculate the workbook to evaluate the formula
            workbook.CalculateFormula();

            // Show the calculated result
            Console.WriteLine("Calculated Value in A1: " + cellA1.Value);

            // Save the workbook (lifecycle save)
            workbook.Save("LocalizedFormulaDemo.xlsx");
        }
    }
}
