// Title: How to recalculate all formulas in an Aspose.Cells workbook after updating cell values using C#
// AI Prompts: Update cells A1 and B1 with numeric values, assign the formula "=A1+B1" to C1, then invoke workbook.CalculateFormula() and read the computed result. | After modifying worksheet data, call Workbook.CalculateFormula() to refresh every dependent formula before saving the workbook.
// Common Searches: C# Aspose.Cells recalculate formulas after changing cell data | Workbook.CalculateFormula example for updating Excel formulas programmatically | How to force formula evaluation in Aspose.Cells before saving workbook | Recalculate dependent formulas in Aspose.Cells after setting cell values | Aspose.Cells .NET recalc all formulas in workbook
// Tags: Aspose.Cells Workbook.CalculateFormula usage | recalculate formulas C# Aspose.Cells | update cell values then evaluate formulas Aspose.Cells | force Excel formula evaluation Aspose.Cells .NET | save workbook after formula recalculation Aspose.Cells

using System;
using Aspose.Cells;

namespace RecalculateFormulasDemo
{
    // Demonstrates creating a workbook, setting numeric values in A1 and B1, adding a formula to C1, invoking Workbook.CalculateFormula() to recalculate all formulas, outputting the result, and saving the workbook as RecalculatedFormulas.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Modify worksheet data
            cells["A1"].PutValue(10);          // Set a numeric value
            cells["B1"].PutValue(20);          // Set another numeric value
            cells["C1"].Formula = "=A1+B1";    // Add a formula that depends on the modified cells

            // Recalculate all formulas in the workbook (rule: CalculateFormula)
            workbook.CalculateFormula();

            // Output the result of the formula to verify calculation
            Console.WriteLine("Result of C1 (A1+B1): " + cells["C1"].Value);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("RecalculatedFormulas.xlsx");
        }
    }
}
