using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a range from A1 to A3 using the CreateRange(string, string) method
        AsposeRange myRange = cells.CreateRange("A1", "A3");
        // Assign a name to the range for later reference in formulas
        myRange.Name = "MyRange";

        // Populate the range with sample numeric values
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);

        // Use the named range in a formula placed in cell B1
        cells["B1"].Formula = "=SUM(MyRange)";

        // Calculate the formula to obtain the result
        workbook.CalculateFormula();

        // Output the calculated sum to the console
        Console.WriteLine("Sum of MyRange (A1:A3) is: " + cells["B1"].IntValue);

        // Save the workbook as an XLSX file
        workbook.Save("CreateRangeDemo.xlsx");
    }
}