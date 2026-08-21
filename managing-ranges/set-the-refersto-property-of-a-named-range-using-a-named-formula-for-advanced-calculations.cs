// Title: Set RefersTo of a Named Range to a Named Formula in Aspose.Cells for .NET
// Description: This example creates a workbook, defines a named formula "Total" that sums cells A1:A3, then creates a second named range "MyRange" whose RefersTo property points to the "Total" formula ("=Total"). The named range is used in cell B1, the workbook evaluates the formulas (result = 60), and the file is saved as NamedFormulaDemo.xlsx.
// Keywords: Aspose.Cells RefersTo property | named formula Aspose.Cells | named range referencing formula .NET | Aspose.Cells calculate formulas | C# workbook named ranges
// Common Searches: Aspose.Cells set RefersTo to another name | how to use a named formula as a range in Aspose.Cells | C# create dependent named ranges Aspose.Cells | reference named formula in worksheet cell Aspose.Cells
// Developer Intent: Create a named range whose RefersTo points to an existing named formula and use it in worksheet calculations.
// Use Cases: Define a reusable total‑sum formula and expose it through a secondary named range for modular design. | Reference a complex calculation from multiple cells via a dependent named range, enabling a single update to propagate. | Expose a calculated value as a named range for external data validation, reporting, or integration within the same workbook.
// AI Prompts: Generate C# code that creates a named formula and assigns another named range's RefersTo to that formula using Aspose.Cells. | Show how to evaluate a named range that references a named formula in an Aspose.Cells worksheet. | Explain how to change a named range's RefersTo at runtime to point to a different named formula in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedFormulaDemo
{
    // This example creates a workbook, defines a named formula "Total" that sums cells A1:A3, then creates a second named range "MyRange" whose RefersTo property points to the "Total" formula ("=Total"). The named range is used in cell B1, the workbook evaluates the formulas (result = 60), and the file is saved as NamedFormulaDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some sample data in column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // -----------------------------------------------------------------
            // 1. Create a named formula "Total" that calculates the sum of A1:A3
            // -----------------------------------------------------------------
            int totalNameIndex = sheet.Workbook.Worksheets.Names.Add("Total");
            Name totalName = sheet.Workbook.Worksheets.Names[totalNameIndex];
            // RefersTo is a formula (starts with '=') that sums the range
            totalName.RefersTo = "=SUM(Sheet1!$A$1:$A$3)";

            // ---------------------------------------------------------------
            // 2. Create another named range "MyRange" that refers to the formula
            //    defined by the name "Total". This demonstrates using a named
            //    formula as the reference of another name.
            // ---------------------------------------------------------------
            int myRangeIndex = sheet.Workbook.Worksheets.Names.Add("MyRange");
            Name myRange = sheet.Workbook.Worksheets.Names[myRangeIndex];
            // The RefersTo property can point to another name by using its name
            myRange.RefersTo = "=Total";

            // ---------------------------------------------------------------
            // 3. Use the named range "MyRange" in a worksheet formula
            // ---------------------------------------------------------------
            sheet.Cells["B1"].Formula = "=MyRange";

            // Calculate all formulas so that B1 shows the result of the sum
            workbook.CalculateFormula();

            // Output the calculated value to the console (optional verification)
            Console.WriteLine("Result of MyRange (should be 60): " + sheet.Cells["B1"].Value);

            // Save the workbook to a file
            workbook.Save("NamedFormulaDemo.xlsx");
        }
    }
}
