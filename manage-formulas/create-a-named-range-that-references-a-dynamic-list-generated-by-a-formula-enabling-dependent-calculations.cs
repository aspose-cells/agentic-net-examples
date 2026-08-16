// Title: Aspose.Cells C# – Create a Dynamic Named Range for a SEQUENCE Spill Array and Use It in Calculations
// Description: Demonstrates how to generate a vertical list with the SEQUENCE function, refresh the spill range, define a workbook named range that points to the spilled array (A1#), and reference that name in a dependent SUM formula using Aspose.Cells for .NET.
// Keywords: Aspose.Cells dynamic named range | C# SEQUENCE spill array | refresh dynamic array formulas | named range from spilled range | dependent calculation Aspose.Cells | Excel dynamic array Aspose | A1# operator Aspose.Cells
// Common Searches: Aspose.Cells create named range for spilled array | C# SEQUENCE function dynamic range Aspose | reference A1# spill range in Aspose.Cells | how to refresh dynamic array formulas Aspose.Cells | use named range in dependent formula Aspose.Cells
// Developer Intent: Define a named range that automatically follows a dynamic SEQUENCE spill range and use it in other formulas.
// Use Cases: Generate a list whose length is driven by a cell value and expose it via a workbook name. | Calculate aggregates (e.g., SUM) on a dynamically sized list without hard‑coding the range. | Refresh dynamic array formulas before calculation to ensure the spill range is up‑to‑date.
// AI Prompts: Write C# code with Aspose.Cells that sets =SEQUENCE(B1) in A1, refreshes the spill, creates a named range "DynamicList" pointing to A1#, and sums the range in C1. | Show how changing the value in B1 automatically resizes the dynamic list and updates the SUM of the named range. | Explain how to create multiple named ranges for different spilled arrays and reference each in separate dependent calculations.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRangeDemo
{
    // Demonstrates how to generate a vertical list with the SEQUENCE function, refresh the spill range, define a workbook named range that points to the spilled array (A1#), and reference that name in a dependent SUM formula using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Prepare data that will drive the dynamic array formula
            // ------------------------------------------------------------
            // Cell B1 will hold the number of rows to generate
            cells["B1"].PutValue(5);

            // ------------------------------------------------------------
            // 2. Set a dynamic array formula that spills based on B1
            // ------------------------------------------------------------
            // The formula =SEQUENCE(B1) creates a vertical list 1..B1
            Cell startCell = cells["A1"];
            startCell.SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

            // Refresh the dynamic array so the spill range is materialized
            workbook.RefreshDynamicArrayFormulas(true);

            // ------------------------------------------------------------
            // 3. Create a named range that points to the spilled range (A1#)
            // ------------------------------------------------------------
            // Add a new name to the workbook's name collection
            int nameIndex = workbook.Worksheets.Names.Add("DynamicList");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];

            // The "#"-operator references the entire spill range of A1
            dynamicName.RefersTo = $"=Sheet1!A1#";

            // ------------------------------------------------------------
            // 4. Use the named range in a dependent calculation
            // ------------------------------------------------------------
            // Example: sum of the dynamic list
            cells["C1"].Formula = "=SUM(DynamicList)";

            // Calculate all formulas (including the dependent one)
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // 5. Save the workbook
            // ------------------------------------------------------------
            workbook.Save("DynamicNamedRangeDemo.xlsx");
        }
    }
}
