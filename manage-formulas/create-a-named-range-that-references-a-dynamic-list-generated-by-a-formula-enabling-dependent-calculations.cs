// Title: Create a Dynamic Named Range from a SEQUENCE Spill and Use It in Calculations – Aspose.Cells for .NET (C#)
// Description: C# sample that generates a variable‑length list with the SEQUENCE function, refreshes the dynamic array spill, creates a workbook‑level named range pointing to the spill range via the A1# operator, and references that name in a dependent SUM formula. The workbook is calculated and saved using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | dynamic array | SEQUENCE function | spill range | A1# operator | named range | refresh dynamic array formulas | dependent formula | SUM | workbook calculation | Excel automation
// Common Searches: Aspose.Cells create named range for dynamic array spill | C# SEQUENCE function spill range Aspose.Cells | how to reference A1# in Aspose.Cells | refresh dynamic array formulas .NET | use named range in dependent formula Aspose.Cells
// Developer Intent: Define a named range that automatically follows a dynamic array spill and use it in other formulas.
// Use Cases: Generate a list whose length is driven by a cell value and aggregate it without hard‑coding addresses. | Share a dynamic list across multiple worksheets via a single named range. | Update the source count (e.g., B1) and have all dependent calculations adjust instantly.
// AI Prompts: Write C# code with Aspose.Cells that sets =SEQUENCE(B1) in A1, refreshes the spill, creates a named range pointing to A1#, and calculates =SUM(DynamicList). | Show how to add a workbook‑level named range for a dynamic array spill (A1#) in Aspose.Cells and use that name in another cell's formula. | Provide an example that changes the value in B1, refreshes dynamic array formulas, and demonstrates the named range automatically expanding or contracting.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRangeDemo
{
    // C# sample that generates a variable‑length list with the SEQUENCE function, refreshes the dynamic array spill, creates a workbook‑level named range pointing to the spill range via the A1# operator, and references that name in a dependent SUM formula. The workbook is calculated and saved using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate data that will drive the dynamic array formula
            //    In this example, cell B1 holds the number of rows to generate
            cells["B1"].PutValue(5); // generate a sequence of 5 numbers

            // 3. Set a dynamic array formula in A1 that spills based on B1
            //    The SEQUENCE function creates a vertical array of numbers 1..B1
            Cell startCell = cells["A1"];
            startCell.SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

            // 4. Refresh dynamic array formulas so the spill range is materialized
            workbook.RefreshDynamicArrayFormulas(true);

            // 5. Create a named range that points to the spilled range (A1#)
            //    The "#" operator references the entire spill range of a dynamic array formula
            int nameIndex = workbook.Worksheets.Names.Add("DynamicList");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = $"=Sheet1!A1#";

            // 6. Use the named range in a dependent calculation (e.g., sum of the list)
            cells["C1"].Formula = "=SUM(DynamicList)";

            // 7. Calculate all formulas so that C1 gets the correct result
            workbook.CalculateFormula();

            // 8. Output the result to the console (optional verification)
            Console.WriteLine($"Sum of dynamic list (should be 15): {cells["C1"].Value}");

            // 9. Save the workbook (uses the standard Aspose.Cells save method)
            workbook.Save("DynamicNamedRangeDemo.xlsx");
        }
    }
}
