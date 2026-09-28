// Title: Ignore #REF! Errors While Calculating Formulas with Aspose.Cells CalculationOptions in C#
// AI Prompts: Enable the IgnoreError flag in CalculationOptions and run Workbook.CalculateFormula so that #REF! cells contribute zero to results. | Show how deleting a column creates a #REF! reference and then evaluate the impacted SUM formula using Aspose.Cells. | Retrieve the value of a cell after recalculating the workbook with error‑ignoring enabled in C#.
// Common Searches: aspocells calculationoptions ignoreerror c# example | how to treat #ref! as zero when calculating Excel formulas in C# | calculate workbook after deleting a column with broken references using Aspose.Cells | c# evaluate SUM formula after column removal with Aspose.Cells | ignore reference errors during formula calculation Aspose.Cells
// Tags: calculationoptions treat errors as zero aspocells c# | manage broken reference in formula evaluation | delete column to generate ref error aspocells | recalculate workbook formulas with error ignoring | sum formula after column removal aspocells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts numeric values, sets a SUM formula, deletes a column to produce a #REF! reference, configures CalculationOptions to ignore errors, recalculates all formulas, and outputs the result where the broken reference is treated as zero.
    public class IgnoreRefErrorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate cells A1 and B1 with numeric values
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);

                // Set a formula in C1 that references both A1 and B1
                cells["C1"].Formula = "=SUM(A1:B1)";

                // Delete column B (index 1) to turn the reference to B1 into #REF!
                worksheet.Cells.DeleteColumn(1);

                // Create calculation options that ignore errors such as #REF!
                CalculationOptions calcOptions = new CalculationOptions
                {
                    IgnoreError = true
                };

                // Calculate all formulas in the workbook using the options
                workbook.CalculateFormula(calcOptions);

                // Output the result of C1. With IgnoreError = true, the #REF! part is treated as zero.
                Console.WriteLine("C1 result after ignoring errors: " + cells["C1"].StringValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            IgnoreRefErrorDemo.Run();
        }
    }
}
