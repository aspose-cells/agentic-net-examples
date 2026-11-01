// Title: Use Aspose.Cells for .NET to apply the UNIQUE dynamic array formula in column D and retrieve deduplicated values
// Common Searches: how to set the UNIQUE dynamic array formula in Aspose.Cells C# | programmatically calculate UNIQUE formula and read results with Aspose.Cells | deduplicate a range using UNIQUE function in a .NET workbook | retrieve distinct values from a column after applying UNIQUE in Aspose.Cells | save workbook after evaluating dynamic array formulas with Aspose.Cells
// Tags: apply UNIQUE formula Aspose.Cells | dynamic array formulas .NET | deduplicate column values Aspose.Cells | calculate workbook formulas programmatically | export Excel file after formula evaluation

using Aspose.Cells;
using System;

// The example creates a new workbook, fills C1:C10 with duplicate strings, assigns the formula =UNIQUE(C1:C10) to D1, forces formula calculation, iterates through column D to output each distinct value, and saves the result as UniqueDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data with duplicates in column C (index 2)
        string[] items = { "Apple", "Banana", "Apple", "Orange", "Banana", "Grape", "Apple", "Grape", "Lemon", "Orange" };
        for (int i = 0; i < items.Length; i++)
        {
            sheet.Cells[i, 2].PutValue(items[i]); // C1:C10
        }

        // Apply the UNIQUE dynamic array formula in D1
        sheet.Cells["D1"].Formula = "=UNIQUE(C1:C10)";

        // Recalculate to evaluate the formula
        workbook.CalculateFormula();

        // Output the deduplicated results from column D
        Console.WriteLine("Unique values from column C:");
        int row = 0;
        while (!string.IsNullOrEmpty(sheet.Cells[row, 3].StringValue)) // D column (index 3)
        {
            Console.WriteLine(sheet.Cells[row, 3].StringValue);
            row++;
        }

        // Save the workbook
        workbook.Save("UniqueDemo.xlsx");
    }
}
