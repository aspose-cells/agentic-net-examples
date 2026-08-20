// Title: C# Aspose.Cells: Create a Dynamic Named Range Using INDEX and a Cell Reference
// Description: Demonstrates how to build a new workbook, fill A1:A10 with values, store a row index in B1, define a named range "DynamicRange" whose RefersTo uses the INDEX function (Sheet1!$A$1:$A$10, Sheet1!$B$1), apply the name in C1, recalculate formulas, output the result, and save the file.
// Keywords: Aspose.Cells dynamic named range | C# INDEX formula named range | RefersTo property Aspose.Cells | cell‑driven named range .NET | calculate formulas Aspose.Cells | Excel dynamic range programmatically | Aspose.Cells example C#
// Common Searches: Aspose.Cells define named range with INDEX | C# dynamic named range based on cell value | How to use RefersTo with INDEX in Aspose.Cells | Update named range when another cell changes .NET | Create Excel dynamic range programmatically
// Developer Intent: Create a named range whose reference is calculated by an INDEX formula that reads its row number from another worksheet cell.
// Use Cases: Return the N‑th item from a list for reporting or validation without hard‑coding addresses. | Link chart data sources to a range that shifts automatically when the user changes an index cell. | Reuse a variable range across multiple formulas, enabling flexible data extraction in dashboards.
// AI Prompts: Generate C# Aspose.Cells code that defines a named range using INDEX with the row index taken from cell B1 and shows the result in C1. | Show how to change the index value in B1, recalculate the workbook, and read the updated value of the dynamic named range. | Provide an example that saves the workbook after creating a cell‑driven dynamic named range with the INDEX function.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRange
{
    // Demonstrates how to build a new workbook, fill A1:A10 with values, store a row index in B1, define a named range "DynamicRange" whose RefersTo uses the INDEX function (Sheet1!$A$1:$A$10, Sheet1!$B$1), apply the name in C1, recalculate formulas, output the result, and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (default name is "Sheet1")
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with sample data (1 to 10)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1:A10
            }

            // Cell B1 will hold the row index that drives the INDEX formula
            // Change this value to see the named range adjust dynamically
            cells["B1"].PutValue(5); // Initially point to the 5th item in column A

            // Create a named range that uses the INDEX function.
            // The formula returns the cell from A1:A10 whose position is given by B1.
            int nameIdx = workbook.Worksheets.Names.Add("DynamicRange");
            Name dynamicName = workbook.Worksheets.Names[nameIdx];
            dynamicName.RefersTo = "=INDEX(Sheet1!$A$1:$A$10, Sheet1!$B$1)";

            // Use the named range in another cell to demonstrate it works
            cells["C1"].Formula = "=DynamicRange";

            // Calculate all formulas so that C1 reflects the current value of B1
            workbook.CalculateFormula();

            // Output the result to the console (optional, for verification)
            Console.WriteLine("Value of DynamicRange (C1): " + cells["C1"].Value);

            // Save the workbook
            workbook.Save("DynamicNamedRange.xlsx");
        }
    }
}
