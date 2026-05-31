using System;
using Aspose.Cells;

namespace AsposeCellsProtectedCalculationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set initial values
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Set a formula that depends on the values above
            cells["B1"].Formula = "=A1+A2";

            // Protect the worksheet with a password
            // ProtectionType.All protects all aspects (cells, objects, etc.)
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Verify that the worksheet is indeed protected
            Console.WriteLine($"Worksheet protected: {sheet.IsProtected}");

            // Calculate formulas after protection
            workbook.CalculateFormula();

            // Display the result of the formula cell
            Console.WriteLine($"B1 formula result after protection: {cells["B1"].IntValue}");

            // Save the workbook (optional, demonstrates that protection persists)
            workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}