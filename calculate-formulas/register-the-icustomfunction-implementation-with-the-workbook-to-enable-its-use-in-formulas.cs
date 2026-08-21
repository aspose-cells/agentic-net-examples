// Title: Register a Custom Add‑In Function (ICustomFunction) in Aspose.Cells Workbook – C# Example
// Description: Demonstrates how to create a Workbook, register a custom Excel add‑in (.xlam) with RegisterAddInFunction, optionally add an alias, use the function in a cell formula, recalculate all formulas, and save the result as an .xlsx file.
// Keywords: Aspose.Cells | C# | .NET | RegisterAddInFunction | custom add‑in | ICustomFunction | Excel custom function | formula calculation | add‑in alias | Workbook
// Common Searches: How to register a custom add‑in function in Aspose.Cells for .NET | RegisterAddInFunction C# example with .xlam file | Create an alias for a custom Excel function using Aspose.Cells | Calculate formulas after adding a custom function in Aspose.Cells | ICustomFunction implementation registration Aspose.Cells
// Developer Intent: Register an ICustomFunction (Excel add‑in) with a workbook so it can be called directly in cell formulas.
// Use Cases: Expose proprietary calculations as a reusable Excel function and invoke it from workbook formulas. | Maintain backward compatibility by assigning an alias to a newly registered custom function. | Register the function once and reuse it across multiple worksheets or workbooks without re‑implementation.
// AI Prompts: Generate C# code that registers a custom add‑in (.xlam) using RegisterAddInFunction and applies it in a worksheet formula with Aspose.Cells. | Explain how to retrieve the function ID from RegisterAddInFunction and create an alias for the same custom function. | Provide error‑handling patterns for missing or invalid add‑in files when registering a custom function in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace CustomFunctionDemo
{
    // Demonstrates registering a custom function (add‑in) with a workbook
    // and using it in cell formulas.
    // Demonstrates how to create a Workbook, register a custom Excel add‑in (.xlam) with RegisterAddInFunction, optionally add an alias, use the function in a cell formula, recalculate all formulas, and save the result as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Path to the add‑in file that contains the custom function implementation.
            // The file must exist; adjust the path as needed for your environment.
            string addInFile = Path.Combine("AddIns", "MyCustomAddIn.xlam");

            // Register the add‑in function with the workbook.
            // Parameters:
            //   addInFile   – the .xlam file containing the function.
            //   "MYFUNC"    – the name of the function as it will be used in formulas.
            //   false       – the path is relative to the workbook, not the add‑in library.
            int functionId = workbook.Worksheets.RegisterAddInFunction(addInFile, "MYFUNC", false);

            // Optionally register an alias for the same function using the returned ID.
            workbook.Worksheets.RegisterAddInFunction(functionId, "MYFUNC_ALIAS");

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data.
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Use the registered custom function in a formula.
            sheet.Cells["B1"].Formula = "=MYFUNC(A1, A2)";

            // Calculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Save the workbook (uses the provided save rule).
            workbook.Save("CustomFunctionDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
