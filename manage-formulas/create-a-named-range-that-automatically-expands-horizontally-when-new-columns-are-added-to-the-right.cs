// Title: C# Aspose.Cells – Create a Horizontally Expanding Dynamic Named Range
// Description: Demonstrates how to define a named range that automatically grows to the right using an INDEX‑COUNTA formula. The example populates A1:C1, adds D1, recalculates formulas, and saves the workbook with the updated range.
// Keywords: Aspose.Cells dynamic named range | C# horizontal expanding range | INDEX COUNTA formula Excel | auto‑grow named range Aspose | .NET Excel named range programmatically | dynamic range chart source
// Common Searches: Aspose.Cells create dynamic named range horizontally | C# expand named range when new columns added | INDEX COUNTA named range example Aspose | auto updating named range Excel .NET | how to make a named range grow with columns
// Developer Intent: Define a named range that automatically includes any new columns added to the right of the original range.
// Use Cases: Keep monthly header cells in a single range so SUM or AVERAGE formulas always cover new months. | Link a chart’s data series to a range that expands as additional period columns are inserted. | Persist a self‑adjusting range in a workbook for downstream reporting or automation scripts.
// AI Prompts: Generate C# code that shifts the dynamic range start from A1 to B1 while retaining horizontal expansion. | Show how to build a vertical dynamic named range that expands downward using Aspose.Cells, including the required formula. | Explain the role of workbook.CalculateFormula() in updating dynamic named ranges and when it can be omitted.

using System;
using Aspose.Cells;

// Demonstrates how to define a named range that automatically grows to the right using an INDEX‑COUNTA formula. The example populates A1:C1, adds D1, recalculates formulas, and saves the workbook with the updated range.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate initial data in the first row (A1, B1, C1)
            cells["A1"].PutValue("Jan");
            cells["B1"].PutValue("Feb");
            cells["C1"].PutValue("Mar");

            // Create a dynamic named range that expands horizontally.
            // The formula uses INDEX together with COUNTA to always point to the last non‑empty cell in row 1.
            // Example: =Sheet1!$A$1:INDEX(Sheet1!$1:$1, COUNTA(Sheet1!$1:$1))
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRange = workbook.Worksheets.Names[nameIndex];
            myRange.RefersTo = "=Sheet1!$A$1:INDEX(Sheet1!$1:$1, COUNTA(Sheet1!$1:$1))";

            // Retrieve the range and display its address before adding a new column
            Aspose.Cells.Range rangeBefore = myRange.GetRange();
            Console.WriteLine("Range before adding column: " + rangeBefore.Address); // Expected: A1:C1

            // Add a new column (D) with data; this should extend the named range automatically
            cells["D1"].PutValue("Apr");

            // Recalculate formulas so that the dynamic named range updates
            workbook.CalculateFormula();

            // Retrieve the range again and display its new address
            Aspose.Cells.Range rangeAfter = myRange.GetRange();
            Console.WriteLine("Range after adding column: " + rangeAfter.Address); // Expected: A1:D1

            // Save the workbook (the named range is persisted)
            string outputPath = "DynamicHorizontalNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
