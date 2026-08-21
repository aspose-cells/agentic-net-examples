// Title: Aspose.Cells for .NET – Create a Dynamic Named Range that Grows with SEQUENCE (C#)
// Description: Demonstrates how to set a SEQUENCE‑based dynamic array in C1, define a named range that points to the spilled range (C1#), refresh the array after changing the row count, and use the name in calculations such as SUM. The workbook is saved as DynamicNamedRangeDemo.xlsx.
// Keywords: Aspose.Cells dynamic named range | C# SEQUENCE formula | spilled array reference | RefreshDynamicArrayFormulas | expand named range automatically | Aspose.Cells .NET example | dynamic array # operator
// Common Searches: Aspose.Cells define named range for spilled array | C# dynamic named range expands with SEQUENCE | Refresh dynamic array after changing source cell Aspose.Cells | How to use C1# reference in Aspose.Cells | Sum values from a dynamic array in Aspose.Cells
// Developer Intent: Create a named range that automatically adjusts to the size of a SEQUENCE‑generated dynamic array and use it in formulas.
// Use Cases: Generate a variable‑length list with SEQUENCE and reference it via a named range for aggregation (SUM, AVERAGE, COUNT). | Change the row‑count cell to enlarge or shrink the spill range, then refresh and recalculate dependent formulas. | Save the workbook containing the dynamic named range for downstream reporting or data export.
// AI Prompts: Show me C# code to create a dynamic named range that points to a spilled SEQUENCE array using Aspose.Cells. | Explain how to refresh dynamic array formulas after updating the source cell in Aspose.Cells for .NET. | Provide examples of using a dynamic named range in SUM, AVERAGE, and COUNT formulas with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to set a SEQUENCE‑based dynamic array in C1, define a named range that points to the spilled range (C1#), refresh the array after changing the row count, and use the name in calculations such as SUM. The workbook is saved as DynamicNamedRangeDemo.xlsx.
class DynamicNamedRangeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // B2 will hold the number of rows for the dynamic array
        cells["B2"].PutValue(5);

        // Set a dynamic array formula in C1 that spills SEQUENCE(B2) rows
        Cell dynamicCell = cells["C1"];
        dynamicCell.SetDynamicArrayFormula("=SEQUENCE(B2)", new FormulaParseOptions(), true);

        // Define a named range that refers to the spilled range (C1#)
        int nameIndex = workbook.Worksheets.Names.Add("MyDynamicRange");
        Name dynamicName = workbook.Worksheets.Names[nameIndex];
        dynamicName.RefersTo = "=Sheet1!C1#";

        // Calculate formulas and refresh dynamic array spill ranges
        workbook.CalculateFormula();
        workbook.RefreshDynamicArrayFormulas(true);

        // Use the dynamic named range in a formula (e.g., SUM)
        cells["E1"].Formula = "=SUM(MyDynamicRange)";
        workbook.CalculateFormula();

        Console.WriteLine("Initial sum of dynamic range: " + cells["E1"].Value);

        // Expand the data by changing B2 (number of rows) to 8
        cells["B2"].PutValue(8);
        workbook.CalculateFormula();
        workbook.RefreshDynamicArrayFormulas(true);
        workbook.CalculateFormula();

        Console.WriteLine("Sum after expanding dynamic range: " + cells["E1"].Value);

        // Save the workbook
        workbook.Save("DynamicNamedRangeDemo.xlsx");
    }
}
