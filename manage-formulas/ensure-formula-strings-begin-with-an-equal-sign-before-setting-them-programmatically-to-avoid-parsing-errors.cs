// Title: Ensure Excel formulas begin with '=' using Aspose.Cells for .NET
// Description: Demonstrates how to programmatically prepend a leading equal sign to formula strings before assigning them to cells with Aspose.Cells for .NET, preventing parsing errors and enabling correct calculation.
// Keywords: Aspose.Cells C# set formula | Excel formula leading equal sign | prepend '=' Aspose.Cells | formula parsing error Aspose | Cell.Formula without = | programmatic formula validation | Excel automation .NET
// Common Searches: add leading equal sign to Excel formula Aspose.Cells | Aspose.Cells formula parsing error missing = | set cell formula C# Aspose.Cells | how to ensure formula starts with = in .NET | validate formula string before assigning Aspose
// Developer Intent: Add a leading '=' to any formula string before setting Cell.Formula to avoid parsing exceptions.
// Use Cases: Sanitize user‑entered formulas before writing them to a worksheet | Generate dynamic formulas from external data sources with guaranteed correct syntax | Batch‑process CSV or database rows containing formulas that may lack the '=' character | Create a reusable helper for consistent formula handling across multiple projects
// AI Prompts: Write a C# utility method for Aspose.Cells that checks a formula string and prepends '=' if missing, then assigns it to a Cell. | Show how to apply the equal‑sign check to an entire range or column in an Aspose.Cells workbook. | Explain why Aspose.Cells requires formulas to start with '=', and provide best practices for avoiding parsing errors when setting formulas programmatically.

using System;
using Aspose.Cells;

// Demonstrates how to programmatically prepend a leading equal sign to formula strings before assigning them to cells with Aspose.Cells for .NET, preventing parsing errors and enabling correct calculation.
class EnsureFormulaDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data that will be used in formulas
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].PutValue(5);

            // Set formulas using the helper that guarantees a leading '='
            SetCellFormula(cells["C1"], "SUM(A1:A2)");   // missing '='
            SetCellFormula(cells["C2"], "=A1*B1");      // already has '='

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the calculated results
            Console.WriteLine("C1 value: " + cells["C1"].Value);
            Console.WriteLine("C2 value: " + cells["C2"].Value);

            // Save the workbook
            workbook.Save("EnsureFormulaDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Helper method that ensures the formula string starts with '=' before setting it
    static void SetCellFormula(Cell cell, string formula)
    {
        // Prepend '=' if it's not already present
        if (!formula.StartsWith("="))
        {
            formula = "=" + formula;
        }

        // Set the formula directly; Aspose.Cells will handle calculation later
        cell.Formula = formula;
    }
}
