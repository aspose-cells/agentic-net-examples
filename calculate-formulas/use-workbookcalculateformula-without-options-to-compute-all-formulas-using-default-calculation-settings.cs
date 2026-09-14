// Title: Calculate all worksheet formulas with default settings using Aspose.Cells Workbook.CalculateFormula in C#
// AI Prompts: Create a new Workbook, assign values and formulas to cells, then call Workbook.CalculateFormula() without parameters to compute every dependent formula. | After the default calculation, read the integer results from cells A1, B1, and C1 and write them to the console. | Optionally persist the workbook to an .xlsx file once the formulas have been evaluated.
// Common Searches: Aspose.Cells C# how to recalculate all formulas in a workbook | Workbook.CalculateFormula default behavior example | evaluate dependent cell formulas with Aspose.Cells without custom settings | C# code to calculate Excel formulas using Aspose.Cells and save the file | default formula calculation in Aspose.Cells workbook
// Tags: calculate formulas Aspose.Cells C# | Workbook.CalculateFormula default settings | evaluate dependent formulas Aspose.Cells | save workbook after calculation Aspose.Cells | default Excel formula evaluation Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, sets a static value and two dependent formulas, runs Workbook.CalculateFormula() with default settings to compute all formulas, prints the results, and saves the workbook as CalculatedWorkbook.xlsx.
    public class WorkbookCalculateFormulaDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet's cells
                Cells cells = workbook.Worksheets[0].Cells;

                // Set initial values and formulas
                cells["A1"].PutValue(5);               // Plain value
                cells["B1"].Formula = "=A1*2";         // Depends on A1
                cells["C1"].Formula = "=B1+10";        // Depends on B1

                // Calculate all formulas using default calculation settings
                workbook.CalculateFormula();

                // Output the calculated results
                Console.WriteLine("A1 value: " + cells["A1"].IntValue);
                Console.WriteLine("B1 value: " + cells["B1"].IntValue);
                Console.WriteLine("C1 value: " + cells["C1"].IntValue);

                // Save the workbook (optional)
                workbook.Save("CalculatedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookCalculateFormulaDemo.Run();
        }
    }
}
