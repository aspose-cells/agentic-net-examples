using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsLocalizedFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook region to Germany to demonstrate German localization
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data that will be used by the formulas
            cells["B1"].PutValue(10);
            cells["C1"].PutValue(20);
            cells["B2"].PutValue(5);
            cells["C2"].PutValue(15);

            // Set a formula using the standard (English) function name
            Cell standardFormulaCell = cells["A1"];
            standardFormulaCell.Formula = "=SUM(B1:C1)";

            // Set a formula using the localized (German) function name via FormulaLocal property
            Cell localizedFormulaCell = cells["A2"];
            localizedFormulaCell.FormulaLocal = "=SUMME(B2:C2)";

            // Display both formulas to verify the difference
            Console.WriteLine("Standard formula (English): " + standardFormulaCell.Formula);
            Console.WriteLine("Localized formula (German): " + localizedFormulaCell.FormulaLocal);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the calculated results
            Console.WriteLine("Result of standard formula (A1): " + standardFormulaCell.Value);
            Console.WriteLine("Result of localized formula (A2): " + localizedFormulaCell.Value);

            // Save the workbook to a file
            workbook.Save("LocalizedFormulaDemo.xlsx");
        }
    }
}