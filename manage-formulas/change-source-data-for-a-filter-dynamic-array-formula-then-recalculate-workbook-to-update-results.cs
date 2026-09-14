// Title: Change the source range of a FILTER dynamic array formula and recalculate the workbook with Aspose.Cells for .NET
// AI Prompts: Update cells referenced by a FILTER dynamic array formula, run workbook.CalculateFormula(), and save the workbook using Aspose.Cells in C#. | Programmatically modify the source data of a FILTER formula, trigger full workbook recalculation, and export the result to a new XLSX file. | Replace values in the FILTER range, invoke Aspose.Cells calculation engine, and write the refreshed workbook to disk.
// Common Searches: Aspose.Cells how to change the range used by a FILTER formula in C# | recalculate FILTER dynamic array after updating source cells with Aspose.Cells | C# example for refreshing FILTER formula results after data modification using Aspose.Cells | update source data for Excel FILTER function and recalc workbook programmatically | Aspose.Cells workbook.CalculateFormula for dynamic array formulas
// Tags: filter formula source range update Aspose.Cells | recalculate workbook formulas Aspose.Cells | modify cell values before CalculateFormula C# | save updated workbook as xlsx Aspose.Cells | dynamic array formula refresh .NET

using Aspose.Cells;

// Loads input.xlsx, changes specific cells in the source range of a FILTER dynamic array formula, recalculates all formulas with workbook.CalculateFormula(), and saves the updated workbook to output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Change source data for the FILTER dynamic array formula -----
        // Example: the FILTER formula uses the range B2:C10 as its source.
        // Update some cells in that source range.
        sheet.Cells["B2"].PutValue(123);   // new value in B2
        sheet.Cells["C5"].PutValue(456);   // new value in C5
        sheet.Cells["B8"].PutValue(789);   // new value in B8
        // Add or modify additional cells as required for your scenario.

        // Recalculate all formulas, including the FILTER dynamic array formula
        workbook.CalculateFormula();

        // Save the updated workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
