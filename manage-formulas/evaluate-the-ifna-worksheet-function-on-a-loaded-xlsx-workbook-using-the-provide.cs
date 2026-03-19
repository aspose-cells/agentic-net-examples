using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put an #N/A error in A1 using the NA() function
        cells["A1"].Formula = "=NA()";

        // Use IFNA to provide a fallback value when A1 is #N/A
        cells["B1"].Formula = "=IFNA(A1, \"NoData\")";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Output the result of the IFNA evaluation
        Console.WriteLine("Result of IFNA in B1: " + cells["B1"].StringValue);
    }
}