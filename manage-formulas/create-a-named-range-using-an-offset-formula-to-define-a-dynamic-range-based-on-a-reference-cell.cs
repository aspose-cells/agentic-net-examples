// Title: C# – Create a Dynamic Named Range with OFFSET in Aspose.Cells
// Description: Demonstrates how to programmatically add a named range that expands automatically using the OFFSET and COUNTA functions, retrieve its address, apply it in a SUM formula, calculate the workbook, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells dynamic named range C# | OFFSET function Aspose.Cells | COUNTA OFFSET named range | programmatic named range Aspose.Cells | SUM formula with named range | calculate formulas Aspose.Cells | save workbook Aspose.Cells
// Common Searches: Aspose.Cells create dynamic named range with OFFSET | C# OFFSET COUNTA named range example | How to use SUM with a named range in Aspose.Cells | Retrieve address of a named range programmatically | Calculate formulas after adding a named range Aspose.Cells
// Developer Intent: Add a self‑adjusting named range using OFFSET/COUNTA, use it in calculations, and persist the workbook.
// Use Cases: Automatically include new rows in column A for totals or charts. | Expose the range address for logging or downstream processing. | Maintain up‑to‑date calculations after data changes without manual range updates.
// AI Prompts: Show how to cap the OFFSET‑based range to a maximum of 10 rows. | Generate code that refreshes the dynamic named range after inserting additional data. | Explain how to bind the dynamic named range to a chart series in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRange
{
    // Demonstrates how to programmatically add a named range that expands automatically using the OFFSET and COUNTA functions, retrieve its address, apply it in a SUM formula, calculate the workbook, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (default name is "Sheet1")
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in column A (A1:A5)
                for (int i = 0; i < 5; i++)
                {
                    cells[i, 0].PutValue(i + 1); // Values 1,2,3,4,5
                }

                // Add a named range that uses OFFSET to create a dynamic range
                // The range starts at A1 and expands downwards based on the count of non‑empty cells in column A
                int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                dynamicName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

                // Retrieve the range via GetRange()
                Aspose.Cells.Range dynamicRange = dynamicName.GetRange();
                Console.WriteLine("Dynamic range address: " + dynamicRange.Address);

                // Use the named range in a formula (e.g., sum of the dynamic range)
                cells["B1"].Formula = "=SUM(DynamicRange)";

                // Calculate formulas so that B1 gets the sum value
                workbook.CalculateFormula();

                // Output the calculated sum to the console
                Console.WriteLine("Sum of dynamic range (should be 15): " + cells["B1"].Value);

                // Save the workbook
                workbook.Save("DynamicNamedRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
