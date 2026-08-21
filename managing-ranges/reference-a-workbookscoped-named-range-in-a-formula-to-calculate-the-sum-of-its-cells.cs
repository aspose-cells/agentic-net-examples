// Title: Aspose.Cells for .NET (C#): Sum a workbook‑scoped named range using a formula
// Description: This example creates a new Workbook, fills cells A1‑A3 with numeric values, defines a workbook‑scoped named range called **MyRange**, inserts the formula **=SUM(MyRange)** into B1, forces calculation with **CalculateFormula()**, prints the result, and saves the file as *NamedRangeSum.xlsx*.
// Keywords: Aspose.Cells | C# | .NET | workbook scoped named range | named range SUM formula | Excel automation | calculate named range total | reference named range in formula | Aspose.Cells example
// Common Searches: Aspose.Cells sum named range C# | how to use workbook scoped named range in Aspose.Cells | reference named range in Excel formula with Aspose.Cells .NET | calculate total of named range using Aspose.Cells | C# Aspose.Cells create and sum named range
// Developer Intent: Create a workbook‑scoped named range and use it in a SUM formula to obtain the total of the referenced cells.
// Use Cases: Aggregate monthly sales figures stored in a column by defining a named range and summing it before exporting the report. | Build financial models where dynamic named ranges are summed to generate subtotals and grand totals automatically. | Design reusable Excel templates that programmatically add named ranges for data validation and then calculate their totals with formulas.
// AI Prompts: Generate C# code that adds a workbook‑scoped named range in Aspose.Cells and calculates its sum with =SUM(). | Show how to set the RefersTo property for a named range and retrieve the computed value after calling CalculateFormula(). | Explain how to reference a workbook‑scoped named range from another worksheet's formula using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace NamedRangeSumExample
{
    // This example creates a new Workbook, fills cells A1‑A3 with numeric values, defines a workbook‑scoped named range called **MyRange**, inserts the formula **=SUM(MyRange)** into B1, forces calculation with **CalculateFormula()**, prints the result, and saves the file as *NamedRangeSum.xlsx*.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (default name is "Sheet1")
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate cells A1, A2, A3 with numeric values
                worksheet.Cells["A1"].PutValue(10);
                worksheet.Cells["A2"].PutValue(20);
                worksheet.Cells["A3"].PutValue(30);

                // Add a workbook‑scoped named range called "MyRange"
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name myRange = workbook.Worksheets.Names[nameIndex];

                // Define the range that the name refers to (absolute reference)
                myRange.RefersTo = "=Sheet1!$A$1:$A$3";

                // Use the named range in a formula to calculate its sum
                worksheet.Cells["B1"].Formula = "=SUM(MyRange)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Output the result of the SUM formula
                Console.WriteLine("Sum of MyRange: " + worksheet.Cells["B1"].Value);

                // Save the workbook to a file
                workbook.Save("NamedRangeSum.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
