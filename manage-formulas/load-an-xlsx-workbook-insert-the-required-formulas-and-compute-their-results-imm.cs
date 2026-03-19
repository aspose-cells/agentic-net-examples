using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook from disk.
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet and its cells collection.
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert values and formulas.
        cells["A1"].PutValue(5);                 // Plain value.
        cells["B1"].Formula = "=A1*2";           // Formula referencing A1.
        cells["C1"].Formula = "=SUM(A1:B1)";     // Formula summing A1 and B1.

        // Calculate all formulas immediately after loading.
        workbook.CalculateFormula();

        // Display the calculated results.
        Console.WriteLine("A1 = " + cells["A1"].Value);
        Console.WriteLine("B1 = " + cells["B1"].Value);
        Console.WriteLine("C1 = " + cells["C1"].Value);

        // Save the workbook with the calculated values (optional).
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}