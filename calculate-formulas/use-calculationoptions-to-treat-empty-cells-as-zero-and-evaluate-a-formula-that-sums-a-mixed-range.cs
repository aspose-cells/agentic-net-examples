// Title: Aspose.Cells C# – Sum a mixed range with blank cells treated as zero using CalculationOptions
// Description: This example creates a workbook, fills a mixed range (A1:C3) with numbers while leaving some cells empty, assigns =SUM(A1:C3) to D1, configures CalculationOptions (TreatEmptyAsZero, IgnoreError, Recursive), evaluates the formula, recalculates the worksheet so D1 reflects the result, and saves the file.
// Keywords: Aspose.Cells CalculationOptions | TreatEmptyAsZero | C# SUM formula blanks | evaluate Excel formula Aspose.Cells | ignore errors recursive calculation | Aspose.Cells sum range with empty cells
// Common Searches: Aspose.Cells treat blank cells as zero | CalculationOptions TreatEmptyAsZero C# | How to sum a range with empty cells in Aspose.Cells | CalculateFormula with IgnoreError and Recursive flags | Save workbook after formula evaluation Aspose.Cells
// Developer Intent: Use CalculationOptions to make blank cells count as zero when evaluating a SUM formula and obtain the result programmatically.
// Use Cases: Calculate total sales where some entries are missing without preprocessing the data. | Generate financial statements that aggregate mixed data ranges while treating blanks as zero. | Re‑calculate worksheets after dynamic updates, ensuring empty cells contribute zero to totals.
// AI Prompts: Show C# code that sets CalculationOptions.TreatEmptyAsZero = true and evaluates =SUM(A1:C3) with Aspose.Cells. | Explain how IgnoreError and Recursive options affect formula evaluation in Aspose.Cells. | Provide a step‑by‑step guide to sum a range containing empty cells and save the workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, fills a mixed range (A1:C3) with numbers while leaving some cells empty, assigns =SUM(A1:C3) to D1, configures CalculationOptions (TreatEmptyAsZero, IgnoreError, Recursive), evaluates the formula, recalculates the worksheet so D1 reflects the result, and saves the file.
    public class TreatEmptyAsZeroDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a mixed range A1:C3 with numbers; some cells remain empty
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            // A3 left empty
            cells["B1"].PutValue(5);
            // B2 left empty
            cells["B3"].PutValue(15);
            cells["C1"].PutValue(0);
            cells["C2"].PutValue(25);
            // C3 left empty

            // Place a formula that sums the whole range
            cells["D1"].Formula = "=SUM(A1:C3)";

            // Additional calculation options (optional)
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true,
                Recursive = true
            };

            // Evaluate the formula directly using the options
            object result = sheet.CalculateFormula("=SUM(A1:C3)", calcOptions);
            Console.WriteLine("Sum with empty cells treated as zero: " + result);

            // Recalculate the whole worksheet so D1 reflects the result
            sheet.CalculateFormula(calcOptions, true);
            Console.WriteLine("Value in D1 after worksheet calculation: " + cells["D1"].Value);

            // Save the workbook
            string outputPath = "TreatEmptyAsZeroDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to " + outputPath);
        }
    }
}
