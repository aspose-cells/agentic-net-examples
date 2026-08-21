// Title: Aspose.Cells C# – Create a Dynamic Named Range that Expands with New Rows in Column O
// Description: Shows how to create a workbook, define a dynamic named range using OFFSET and COUNTA that automatically grows as rows are appended to column O, apply the range in a SUM formula, recalculate, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | dynamic named range | OFFSET | COUNTA | auto‑grow range | column O | SUM formula | Workbook automation | Worksheet programming
// Common Searches: Aspose.Cells create dynamic named range | C# OFFSET COUNTA named range | auto expanding range Aspose.Cells | sum dynamic column O Aspose | recalculate formulas after adding rows Aspose.Cells
// Developer Intent: Create a programmatic named range that automatically expands when new rows are added to column O and use it in calculations.
// Use Cases: Generate reports where the size of a data column changes over time. | Build charts that reference a range that grows with incoming data. | Calculate running totals on a column without manually updating the range. | Distribute a workbook to downstream users while keeping formulas resilient to added rows.
// AI Prompts: Write C# code with Aspose.Cells to define a named range that uses OFFSET and COUNTA to cover all non‑empty cells in column O. | Show how to append rows to column O, recalculate formulas, and retrieve the updated SUM of a dynamic named range. | Explain how to reference the dynamic named range in charts or other worksheets using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRange
{
    // Shows how to create a workbook, define a dynamic named range using OFFSET and COUNTA that automatically grows as rows are appended to column O, apply the range in a SUM formula, recalculate, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (default name is "Sheet1")
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some initial data in column O (index 14, zero‑based)
            sheet.Cells["O1"].PutValue("Header");
            sheet.Cells["O2"].PutValue(10);
            sheet.Cells["O3"].PutValue(20);
            sheet.Cells["O4"].PutValue(30);

            // Add a named range that expands automatically as rows are added to column O.
            // The formula uses OFFSET together with COUNTA to determine the current height.
            // =OFFSET(Sheet1!$O$2,0,0,COUNTA(Sheet1!$O:$O)-1,1)
            // - Starts at O2 (first data row)
            // - Height = number of non‑empty cells in column O minus the header row
            int nameIndex = workbook.Worksheets.Names.Add("DynamicO");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$O$2,0,0,COUNTA({sheet.Name}!$O:$O)-1,1)";

            // Demonstrate that the named range works by using it in a formula
            // Sum of all values in the dynamic range
            sheet.Cells["P1"].Formula = "=SUM(DynamicO)";

            // Calculate formulas so that the sum is evaluated
            workbook.CalculateFormula();

            Console.WriteLine($"Initial sum in P1: {sheet.Cells["P1"].Value}");

            // Add more rows to column O
            sheet.Cells["O5"].PutValue(40);
            sheet.Cells["O6"].PutValue(50);

            // Re‑calculate to reflect the new rows
            workbook.CalculateFormula();

            Console.WriteLine($"Updated sum in P1 after adding rows: {sheet.Cells["P1"].Value}");

            // Save the workbook
            workbook.Save("DynamicNamedRangeDemo.xlsx");
        }
    }
}
