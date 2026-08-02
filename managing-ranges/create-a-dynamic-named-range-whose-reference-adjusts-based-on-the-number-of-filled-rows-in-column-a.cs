// Title: C# – Create a dynamic named range that expands with filled rows in column A using Aspose.Cells
// Description: Demonstrates how to generate a workbook, populate A1:A5, add a named range "MyDynamicRange" defined by an OFFSET‑COUNTA formula, use it in a SUM formula (B1), calculate the result, output 15, and save the file as DynamicNamedRange.xlsx.
// Keywords: Aspose.Cells | C# | .NET | dynamic named range | OFFSET formula | COUNTA function | auto‑expanding range | Excel named range | sum formula | workbook calculation
// Common Searches: Aspose.Cells create dynamic named range | C# OFFSET COUNTA named range example | auto expanding range column A Aspose.Cells | how to use named range in Aspose.Cells formula | calculate sum of dynamic range with Aspose.Cells
// Developer Intent: Add a named range that automatically adjusts its size to the number of non‑empty rows in column A.
// Use Cases: Define a self‑adjusting range for column A and reference it in aggregate formulas such as SUM or AVERAGE. | Reuse the same dynamic range in lookup functions (e.g., VLOOKUP, MATCH) without hard‑coding the row count. | Link the dynamic range to chart data sources so the chart expands as new rows are added.
// AI Prompts: Write C# code with Aspose.Cells that creates a dynamic named range based on non‑empty cells in column A and uses it in a SUM formula. | Show how to modify the OFFSET/CounTA expression to start from B2 and include columns B‑C. | Provide an example of binding a chart series to the dynamic named range created with Aspose.Cells.

using System;
using Aspose.Cells;

namespace DynamicNamedRangeDemo
{
    // Demonstrates how to generate a workbook, populate A1:A5, add a named range "MyDynamicRange" defined by an OFFSET‑COUNTA formula, use it in a SUM formula (B1), calculate the result, output 15, and save the file as DynamicNamedRange.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with sample data (filled rows)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1:A5 = 1,2,3,4,5
            }

            // Add a dynamic named range that expands based on the number of non‑empty rows in column A
            // Formula: =OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)
            int nameIdx = workbook.Worksheets.Names.Add("MyDynamicRange");
            Name dynamicName = workbook.Worksheets.Names[nameIdx];
            dynamicName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

            // Use the dynamic named range in a formula (e.g., sum of the range)
            cells["B1"].Formula = "=SUM(MyDynamicRange)";

            // Calculate formulas to obtain the result
            workbook.CalculateFormula();

            // Output the calculated sum (should be 1+2+3+4+5 = 15)
            Console.WriteLine("Sum of dynamic range: " + cells["B1"].Value);

            // Save the workbook
            workbook.Save("DynamicNamedRange.xlsx");
        }
    }
}
