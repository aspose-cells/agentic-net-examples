using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data and formulas
        cells["A1"].PutValue(5);               // Simple value
        cells["B1"].Formula = "=A1*2";          // Formula referencing A1
        cells["C1"].Formula = "=B1+10";         // Formula referencing B1

        // Calculate all formulas in the workbook (core operation)
        workbook.CalculateFormula();

        // Retrieve and display the updated cell values after calculation
        Console.WriteLine("A1 value: " + cells["A1"].Value);
        Console.WriteLine("B1 value (A1*2): " + cells["B1"].Value);
        Console.WriteLine("C1 value (B1+10): " + cells["C1"].Value);

        // Save the workbook to a file (lifecycle: save)
        workbook.Save("CalculatedWorkbook.xlsx");
    }
}