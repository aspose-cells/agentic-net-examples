using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put a numeric value in A1
        cells["A1"].PutValue(10);

        // Set the IFNA formula in B1: if A1 is not an error, return its value; otherwise return "fallback"
        cells["B1"].Formula = "=IFNA(A1, \"fallback\")";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Display the result of the IFNA formula
        Console.WriteLine("B1 result: " + cells["B1"].StringValue);
    }
}