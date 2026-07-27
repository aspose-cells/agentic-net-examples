using System;
using Aspose.Cells;

namespace AsposeCellsProtectedFormulaDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with values and a formula that depends on them
            cells["A1"].PutValue(5);               // Input value
            cells["B1"].Formula = "=A1*2";          // Formula to calculate
            cells["C1"].Formula = "=B1+10";         // Dependent formula

            // Protect the worksheet with a password
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Verify protection status (optional)
            Console.WriteLine($"Worksheet protected: {sheet.IsProtected}");

            // Calculate all formulas in the workbook after protection
            workbook.CalculateFormula();

            // Display the calculated results
            Console.WriteLine($"A1 value: {cells["A1"].IntValue}");
            Console.WriteLine($"B1 formula result: {cells["B1"].IntValue}");
            Console.WriteLine($"C1 formula result: {cells["C1"].IntValue}");

            // Save the workbook (optional)
            workbook.Save("ProtectedFormulaDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}