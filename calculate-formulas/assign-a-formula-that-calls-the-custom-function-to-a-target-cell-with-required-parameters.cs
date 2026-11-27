// Title: Register an XLA/XLAM add‑in function and assign its formula to a cell using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# register xlam add‑in function and use it in a worksheet formula | how to assign a custom UDF formula to a cell with Aspose.Cells .NET | example of calling a custom function from an Excel add‑in in Aspose.Cells | calculate result of a custom function after setting its formula in Aspose.Cells | save workbook after invoking a custom add‑in function with Aspose.Cells C#
// Tags: add‑in function registration Aspose.Cells | custom UDF formula assignment Aspose.Cells | invoke custom function with cell references Aspose.Cells | force formula calculation Aspose.Cells | save workbook with custom function result Aspose.Cells | load XLA add‑in file Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads a .xlam add‑in, registers its custom function with the workbook, writes sample values to A1‑A3, sets the formula =MY_UDF(A1, A2, A3) in B1, forces calculation, outputs the result, and saves the workbook as CustomFunctionDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Locate the add‑in file (XLA/XLAM) that contains the custom UDF.
            //    The file must exist on disk; otherwise the demo cannot proceed.
            // -----------------------------------------------------------------
            string addInFileName = "MyCustomUDF.xlam";
            string addInPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, addInFileName);

            if (!File.Exists(addInPath))
            {
                Console.WriteLine($"Error: Add‑in file not found at '{addInPath}'.");
                Console.WriteLine("Please copy a valid .xlam file that defines the custom function into the executable folder.");
                return;
            }

            // -----------------------------------------------------------------
            // 2. Name of the custom function as defined inside the add‑in.
            //    Example: function MY_UDF(arg1, arg2, arg3) { … }
            // -----------------------------------------------------------------
            string customFunctionName = "MY_UDF";

            // -----------------------------------------------------------------
            // 3. Create a new workbook.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // 4. Register the add‑in function with the workbook.
            //    RegisterAddInFunction returns an integer ID that can be reused
            //    for additional functions from the same add‑in file.
            // -----------------------------------------------------------------
            int functionId = workbook.Worksheets.RegisterAddInFunction(addInPath, customFunctionName, false);
            Console.WriteLine($"Registered add‑in function '{customFunctionName}' with ID {functionId}.");

            // -----------------------------------------------------------------
            // 5. (Optional) Register another function from the same add‑in using
            //    the ID obtained above.
            // -----------------------------------------------------------------
            // workbook.Worksheets.RegisterAddInFunction(functionId, "ANOTHER_UDF");

            // -----------------------------------------------------------------
            // 6. Access the first worksheet and its Cells collection.
            // -----------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -----------------------------------------------------------------
            // 7. Populate some sample data that the custom function will consume.
            // -----------------------------------------------------------------
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // -----------------------------------------------------------------
            // 8. Assign a formula that calls the custom function with required parameters.
            //    The formula string must start with '=' and use commas as argument separators.
            //    Example formula: =MY_UDF(A1, A2, A3)
            // -----------------------------------------------------------------
            cells["B1"].Formula = $"={customFunctionName}(A1, A2, A3)";

            // -----------------------------------------------------------------
            // 9. Force calculation of all formulas, including the custom one.
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            // -----------------------------------------------------------------
            // 10. Retrieve and display the result of the custom function.
            // -----------------------------------------------------------------
            Console.WriteLine($"Result of custom function in B1: {cells["B1"].Value}");

            // -----------------------------------------------------------------
            // 11. Save the workbook to verify that the formula and result are persisted.
            // -----------------------------------------------------------------
            string outputPath = "CustomFunctionDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
