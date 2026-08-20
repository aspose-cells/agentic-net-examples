// Title: Aspose.Cells .NET: Create a Horizontally Expanding Dynamic Named Range with OFFSET & COUNTA
// Description: Demonstrates how to programmatically add a named range that automatically widens as new columns are inserted. The example builds a workbook, defines the range using OFFSET and COUNTA, sums the range, extends the table, recalculates formulas, and saves the file.
// Keywords: Aspose.Cells | .NET | C# | dynamic named range | horizontal expansion | OFFSET function | COUNTA function | auto‑update range | Excel named range programmatically | calculate formulas | expand columns | named range formula
// Common Searches: Aspose.Cells create dynamic named range | C# named range that expands with new columns | OFFSET COUNTA named range Aspose.Cells | auto updating named range in Excel using code | sum dynamic horizontal range Aspose.Cells
// Developer Intent: Define a named range that automatically expands horizontally when columns are added and use it in calculations.
// Use Cases: Maintain a running total that adjusts as the data table grows to the right. | Supply a chart data source that updates when new metric columns are appended. | Apply conditional formatting to a range that expands with additional columns.
// AI Prompts: Generate C# code with Aspose.Cells to create a named range that expands horizontally based on non‑empty cells in the first row. | Show how to recalculate formulas after adding columns and retrieve the updated sum from the dynamic range. | Explain how to convert the dynamic named range into an Aspose.Cells Table for further data manipulation.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeDynamic
{
    // Demonstrates how to programmatically add a named range that automatically widens as new columns are inserted. The example builds a workbook, defines the range using OFFSET and COUNTA, sums the range, extends the table, recalculates formulas, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some initial data in the first row (horizontal table)
            // The table will start at A1 and may expand to the right
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["C2"].PutValue(30);

            // Create a named range that expands horizontally with the table.
            // The formula uses OFFSET together with COUNTA to count the number of filled cells in row 1.
            // It always starts at A1, has 1 row height, and a width equal to the number of non‑empty cells in row 1.
            int nameIndex = workbook.Worksheets.Names.Add("DynamicRow");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = "=OFFSET(Data!$A$1,0,0,1,COUNTA(Data!$1:$1))";

            // Demonstrate that the named range works in a formula.
            // Sum the values of the dynamic range (which currently includes A2:C2).
            sheet.Cells["E1"].Formula = "=SUM(DynamicRow)";

            // Calculate formulas so that the sum is evaluated.
            workbook.CalculateFormula();

            // Output the result to the console.
            Console.WriteLine("Sum of dynamic range: " + sheet.Cells["E1"].Value);

            // Expand the table horizontally by adding two more columns.
            sheet.Cells["D1"].PutValue("Header4");
            sheet.Cells["E1"].PutValue("Header5");
            sheet.Cells["D2"].PutValue(40);
            sheet.Cells["E2"].PutValue(50);

            // Re‑calculate to reflect the expanded range.
            workbook.CalculateFormula();

            // Output the updated sum.
            Console.WriteLine("After expanding, sum of dynamic range: " + sheet.Cells["E1"].Value);

            // Save the workbook.
            workbook.Save("DynamicNamedRange.xlsx");
        }
    }
}
