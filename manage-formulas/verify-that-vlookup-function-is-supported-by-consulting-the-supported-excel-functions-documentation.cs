using System;
using Aspose.Cells;

class VerifyVlookupSupport
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Build a simple lookup table
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Price");
        cells["A2"].PutValue("Apple");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("Banana");
        cells["B3"].PutValue(20);

        // Apply VLOOKUP formula that should return the price of "Apple"
        cells["D1"].Formula = "=VLOOKUP(\"Apple\",A1:B3,2,FALSE)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // If the function is unsupported, Aspose.Cells marks it as a custom function
        bool hasCustomFunction = cells["D1"].HasCustomFunction;

        // VLOOKUP is supported when HasCustomFunction is false
        Console.WriteLine("Is VLOOKUP supported? " + (!hasCustomFunction));

        // Display the result of the VLOOKUP formula
        Console.WriteLine("VLOOKUP result: " + cells["D1"].StringValue);
    }
}