// Title: Calculate a SUMPRODUCT formula with Aspose.Cells for .NET to sum values for a specific category in Excel
// Common Searches: how to calculate conditional sum with SUMPRODUCT using Aspose.Cells C# | set array-like formula in Aspose.Cells to sum values for Group A | programmatically add SUMPRODUCT formula to an Excel workbook with Aspose.Cells .NET | aggregate grouped numeric data in Excel via Aspose.Cells formula calculation | calculate all workbook formulas after populating data using Aspose.Cells C#
// Tags: SUMPRODUCT implementation Aspose.Cells .NET | assign Excel formula via C# Aspose.Cells | calculate workbook formulas programmatically | populate worksheet data for aggregation | grouped value aggregation Excel C#

using System;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaDemo
{
    // The example creates a new workbook, fills column A with category names and column B with numeric values, assigns a SUMPRODUCT formula to cell C2 that sums the values where the category equals "Group A", triggers formula calculation, outputs the result, and saves the file as ArrayFormulaDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();                     // creates a new empty workbook
                Worksheet sheet = workbook.Worksheets[0];              // get the first (default) worksheet
                Cells cells = sheet.Cells;                             // shortcut to the cells collection

                // -----------------------------------------------------------------
                // 2. Populate sample data.
                //    Column A : Category (Group A / Group B)
                //    Column B : Numeric values to be aggregated.
                // -----------------------------------------------------------------
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                string[] categories = { "Group A", "Group B", "Group A", "Group B", "Group A" };
                double[] values = { 100, 150, 200, 250, 300 };

                for (int i = 0; i < categories.Length; i++)
                {
                    cells[i + 2, 0].PutValue(categories[i]);   // A2, A3, ...
                    cells[i + 2, 1].PutValue(values[i]);      // B2, B3, ...
                }

                // -----------------------------------------------------------------
                // 3. Define a formula that aggregates (sums) the values belonging to "Group A".
                //    Using SUMPRODUCT avoids the need for an explicit array formula.
                //    Excel equivalent: =SUMPRODUCT((A2:A6="Group A")*B2:B6)
                // -----------------------------------------------------------------
                string formula = "SUMPRODUCT((A2:A6=\"Group A\")*B2:B6)";
                cells["C2"].Formula = formula;   // set regular formula (no FormulaArray property in this version)

                // -----------------------------------------------------------------
                // 4. Calculate all formulas in the workbook.
                // -----------------------------------------------------------------
                workbook.CalculateFormula();

                // -----------------------------------------------------------------
                // 5. Retrieve and display the result of the formula.
                // -----------------------------------------------------------------
                double result = cells["C2"].DoubleValue;
                Console.WriteLine("Formula result (sum of Group A values): " + result);

                // -----------------------------------------------------------------
                // 6. Save the workbook to verify the formula and result.
                // -----------------------------------------------------------------
                string outputPath = "ArrayFormulaDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                // Any unexpected error will be printed to the console.
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
