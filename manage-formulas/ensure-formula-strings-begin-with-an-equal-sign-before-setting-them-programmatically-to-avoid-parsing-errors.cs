// Title: Add Missing '=' to Excel Formulas in Aspose.Cells for .NET
// Description: Shows a C# helper that guarantees every formula string begins with an equal sign before assigning it to a cell via the Formula property, avoiding parsing errors. The sample creates a workbook, writes values, corrects raw formulas, calculates results, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | Excel formula validation | prepend equal sign | Formula property | parsing error prevention | EnsureFormula helper | workbook calculation | Excel automation | programmatic formula correction
// Common Searches: how to add '=' to Excel formulas using Aspose.Cells C# | Aspose.Cells formula parsing error missing equal sign | C# helper to ensure Excel formula starts with = | set formula property in Aspose.Cells without error | validate raw formula strings before assigning to cells
// Developer Intent: Validate and automatically prepend an equal sign to any formula string before setting Cells[i].Formula in Aspose.Cells to prevent runtime parsing exceptions.
// Use Cases: Sanitize user‑entered or external system‑generated formulas before inserting them into a workbook. | Standardize formula syntax across multiple worksheets in large‑scale reporting pipelines. | Create a reusable utility that safely assigns formulas in automated Excel generation tasks.
// AI Prompts: Generate a C# method for Aspose.Cells that adds a leading '=' to a formula only if it is missing, and demonstrate its use with sample cells. | Write C# code that scans a range of cells in an Aspose.Cells workbook, ensures each cell's Formula property starts with '=', and handles empty strings gracefully. | Provide an example that corrects formulas with a helper, saves the workbook, then reopens it to verify that all formulas contain the leading '='.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaHelperDemo
{
    // Shows a C# helper that guarantees every formula string begins with an equal sign before assigning it to a cell via the Formula property, avoiding parsing errors. The sample creates a workbook, writes values, corrects raw formulas, calculates results, and saves the file.
    class Program
    {
        // Ensures that a formula string starts with '='.
        // If the input already starts with '=', it is returned unchanged.
        // Otherwise, '=' is prepended.
        static string EnsureFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula))
                return formula; // Let the caller handle empty strings.

            return formula.StartsWith("=") ? formula : "=" + formula;
        }

        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Put some sample values that will be used in formulas.
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["B1"].PutValue(5);

                // Example formulas without leading '='.
                string rawFormula1 = "SUM(A1,A2)";          // Should become "=SUM(A1,A2)"
                string rawFormula2 = "A1*B1";              // Should become "=A1*B1"
                string rawFormula3 = "=AVERAGE(A1:A2)";    // Already has '=', keep as is.

                // Set formulas using the Formula property (automatically handles parsing).
                cells["C1"].Formula = EnsureFormula(rawFormula1);
                cells["C2"].Formula = EnsureFormula(rawFormula2);
                cells["C3"].Formula = EnsureFormula(rawFormula3);

                // Calculate all formulas in the workbook.
                workbook.CalculateFormula();

                // Output the results to verify correct calculation.
                Console.WriteLine("C1 (SUM): " + cells["C1"].Value);          // Expected 30
                Console.WriteLine("C2 (Product): " + cells["C2"].Value);    // Expected 50
                Console.WriteLine("C3 (Average): " + cells["C3"].Value);    // Expected 15

                // Optionally, save the workbook to inspect the formulas.
                workbook.Save("FormulaHelperDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
