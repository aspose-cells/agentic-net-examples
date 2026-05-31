using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample operations that affect cell E3
        cells["A1"].PutValue(5);                     // A1 = 5
        cells["B1"].PutValue(10);                    // B1 = 10
        cells["C1"].Formula = "=A1+B1";              // C1 = 15
        cells["D1"].Formula = "=C1*2";               // D1 = 30
        cells["E3"].Formula = "=D1+100";             // E3 = 130

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the final calculated value of E3
        object e3Value = cells["E3"].Value;

        // Display the result
        Console.WriteLine("Final calculated value of E3: " + e3Value);
    }
}