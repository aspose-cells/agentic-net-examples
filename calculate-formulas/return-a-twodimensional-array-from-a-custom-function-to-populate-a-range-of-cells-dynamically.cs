// Title: Insert a pre‑calculated 3×2 two‑dimensional array into an Excel range using SetDynamicArrayFormula in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a 3 × 2 object[][] and writes it to cell A1 with SetDynamicArrayFormula while turning off automatic calculation. | Show how to supply a placeholder custom function (e.g., =MYFUNC()) and pass the pre‑computed array values to a dynamic‑array cell in Aspose.Cells. | Explain how to refresh dynamic‑array formulas after inserting the values and then save the workbook.
// Common Searches: Aspose.Cells how to use SetDynamicArrayFormula with a pre‑filled object[][] array | C# write a two‑dimensional array to an Excel range without recalculation | populate Excel dynamic array from custom function result Aspose.Cells .NET | refresh dynamic array formulas after manual value insertion Aspose.Cells
// Tags: pre‑calculated object[][] dynamic array insertion | user‑defined function returning 2D array Aspose.Cells | prevent recalculation of dynamic array .NET | refresh dynamic array formulas workbook | populate Excel range from C# object array

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayDemo
{
    // // Demonstrates creating a 3×2 object[][], inserting it into cell A1 via SetDynamicArrayFormula with calculation disabled, optionally refreshing dynamic arrays, and saving the workbook as DynamicArrayResult.xlsx.
    class Program
    {
        // This method simulates a custom function that returns a two‑dimensional array.
        // In a real scenario you could register a custom function with Aspose.Cells,
        // but for simplicity we calculate the values here and pass them to the
        // SetDynamicArrayFormula overload that accepts pre‑calculated values.
        static object[][] GetSampleArray()
        {
            // Create a 3 × 2 array.
            object[][] result = new object[3][];
            result[0] = new object[] { 10, 20 };
            result[1] = new object[] { 30, 40 };
            result[2] = new object[] { 50, 60 };
            return result;
        }

        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Prepare the data that the custom function would return.
            object[][] arrayValues = GetSampleArray();

            // 3. Set a dynamic array formula in cell A1.
            // The formula string can be any valid Excel formula; here we use a placeholder
            // function name "MYFUNC". Because we supply the pre‑calculated values, the
            // formula itself will not be evaluated – the values will be written directly.
            Cell targetCell = cells["A1"];
            string formula = "=MYFUNC()";

            // SetDynamicArrayFormula overload:
            // (formula, parseOptions, values, calculateRange, calculateValue)
            // - calculateRange = false  -> use the dimensions of the supplied values.
            // - calculateValue = false  -> do not recalculate; use supplied values.
            targetCell.SetDynamicArrayFormula(
                formula,
                new FormulaParseOptions(),
                arrayValues,
                calculateRange: false,
                calculateValue: false);

            // 4. (Optional) Refresh dynamic array formulas if you need Excel to recalculate
            // any other dynamic arrays in the workbook.
            workbook.RefreshDynamicArrayFormulas(false);

            // 5. Save the workbook.
            workbook.Save("DynamicArrayResult.xlsx");
        }
    }
}
