using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Demonstrates how to register a custom function (implemented as an add‑in) with a workbook
    // and then use that function in cell formulas.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (uses the standard create rule)
            Workbook workbook = new Workbook();

            // 2. Register the custom function.
            //    The function is assumed to be defined in an Excel add‑in file (XLA/XLAM).
            //    The third parameter indicates that the path is relative to the workbook,
            //    not to the Aspose.Cells add‑in library.
            string addInPath = Path.Combine("AddIns", "MyCustomFunctions.xlam"); // adjust as needed
            string functionName = "MY_UDF"; // the name of the function defined in the add‑in
            workbook.Worksheets.RegisterAddInFunction(addInPath, functionName, false);

            // 3. Use the registered function in a cell formula.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].Formula = $"={functionName}(A1, A2)";

            // 4. Calculate the workbook so the custom function is evaluated.
            //    If the add‑in implements the function correctly, the result will appear in B1.
            workbook.CalculateFormula();

            // 5. Output the result to the console (optional verification).
            Console.WriteLine($"Result of {functionName}(A1, A2): {sheet.Cells["B1"].Value}");

            // 6. Save the workbook (uses the standard save rule)
            workbook.Save("CustomFunctionDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}