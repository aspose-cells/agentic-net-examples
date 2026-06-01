using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in cells A1:A3
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);

        // Create a named range called "MyRange" that refers to A1:A3
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        Name namedRange = workbook.Worksheets.Names[nameIndex];
        namedRange.RefersTo = $"={sheet.Name}!$A$1:$A$3";

        // Set a formula in B1 that sums the named range
        cells["B1"].Formula = "=SUM(MyRange)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve and display the calculated result
        double sumResult = cells["B1"].DoubleValue;
        Console.WriteLine($"Sum of MyRange (A1:A3): {sumResult}");

        // Verify that the result matches the expected value (10+20+30 = 60)
        if (Math.Abs(sumResult - 60) < 0.0001)
        {
            Console.WriteLine("Verification passed.");
        }
        else
        {
            Console.WriteLine("Verification failed.");
        }

        // (Optional) Save the workbook if you want to inspect it
        // workbook.Save("NamedRangeSum.xlsx");
    }
}