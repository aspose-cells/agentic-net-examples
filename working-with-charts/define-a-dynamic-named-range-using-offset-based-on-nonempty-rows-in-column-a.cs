// Title: Create a dynamic named range with OFFSET and COUNTA for non‑empty rows in column A using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that defines a named range using OFFSET and COUNTA to automatically include all populated cells in column A. | Show how to reference the dynamic named range in a SUM formula, trigger formula calculation, and retrieve the computed value. | Demonstrate saving the workbook after the dynamic range is created and the formula is evaluated.
// Common Searches: asp.net aspose.cells create dynamic named range with offset based on column A values | c# use COUNTA in named range formula with Aspose.Cells | calculate sum of a dynamic range defined by OFFSET in Aspose.Cells .NET | save workbook after defining dynamic named range in Aspose.Cells C# example
// Tags: OFFSET dynamic named range Aspose.Cells | COUNTA based range definition C# | SUM formula using named range Aspose.Cells | calculate workbook formulas Aspose.Cells | save Excel workbook with dynamic range Aspose.Cells

using System;
using Aspose.Cells;

namespace DynamicNamedRangeDemo
{
    // The example creates a new workbook, fills column A with sample data, adds a named range called DynamicRange using an OFFSET formula that counts non‑empty rows via COUNTA, inserts a SUM(DynamicRange) formula in cell B1, calculates all formulas, prints the sum, and saves the file as DynamicNamedRangeDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate column A with sample data (some non‑empty rows)
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(20);
            cells["A4"].PutValue(30);
            // Row 5 left blank to demonstrate COUNTA counting only non‑empty rows

            // 3. Define a dynamic named range using OFFSET and COUNTA
            //    Formula: =OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)
            int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$A$1,0,0,COUNTA({sheet.Name}!$A:$A),1)";

            // 4. Use the named range in a formula (e.g., sum of the range)
            cells["B1"].Formula = "=SUM(DynamicRange)";

            // 5. Calculate formulas so that the result is materialized
            workbook.CalculateFormula();

            // 6. Output the calculated sum to console (optional verification)
            Console.WriteLine("Sum of DynamicRange: " + cells["B1"].Value);

            // 7. Save the workbook (lifecycle: save)
            workbook.Save("DynamicNamedRangeDemo.xlsx");
        }
    }
}
