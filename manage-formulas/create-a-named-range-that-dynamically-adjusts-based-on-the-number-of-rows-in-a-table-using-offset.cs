// Title: C# Aspose.Cells Sample: Create a Dynamic Named Range Using OFFSET and COUNTA
// Description: This .NET example builds a workbook, fills column A with a header and data rows, defines a named range that automatically expands via OFFSET‑COUNTA, uses it in a SUM formula, recalculates the sheet, and saves the file.
// Keywords: Aspose.Cells | C# | dynamic named range | OFFSET function | COUNTA | Excel automation | named range formula | auto‑expand range | workbook calculation | sample code
// Common Searches: Aspose.Cells create expanding named range | OFFSET COUNTA C# example | dynamic range in Excel using Aspose | sum over auto‑growing range Aspose.Cells | recalculate formulas after adding named range .NET
// Developer Intent: Define a workbook‑level named range that resizes automatically as rows are added to a column.
// Use Cases: Compute totals for a column whose length varies over time | Drive charts or pivot tables that need to reflect newly added rows | Apply data validation, conditional formatting, or formulas without manually updating the address
// AI Prompts: Write C# code that creates a dynamic range for multiple adjacent columns using OFFSET | Show how to modify the OFFSET formula to ignore blank cells at the bottom of the range | Explain how to trigger a full workbook recalculation after inserting rows that affect an expanding range

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRange
{
    // This .NET example builds a workbook, fills column A with a header and data rows, defines a named range that automatically expands via OFFSET‑COUNTA, uses it in a SUM formula, recalculates the sheet, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (A1 is header)
            cells["A1"].PutValue("Item");
            for (int i = 2; i <= 10; i++)          // 9 data rows
            {
                cells[$"A{i}"].PutValue($"Value {i - 1}");
            }

            // Add a named range that expands dynamically with the number of rows in column A
            // OFFSET(start, rows, cols, height, width)
            // start: A2 (first data cell)
            // rows: 0, cols: 0 (no offset)
            // height: COUNTA(A:A)-1 (total non‑empty rows minus header)
            // width: 1 (single column)
            int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$A$2,0,0,COUNTA({sheet.Name}!$A:$A)-1,1)";

            // Optional: demonstrate usage of the named range in a formula
            cells["B1"].Formula = "=SUM(DynamicRange)";

            // Calculate formulas so that B1 shows the sum of the dynamic range
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("DynamicRangeDemo.xlsx");
        }
    }
}
