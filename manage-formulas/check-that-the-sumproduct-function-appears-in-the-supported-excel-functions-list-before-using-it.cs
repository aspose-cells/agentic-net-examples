// Title: Check SUMPRODUCT support in Aspose.Cells for .NET before applying the formula
// Description: Creates a workbook, adds sample data, and uses a temporary cell to evaluate the SUMPRODUCT formula. By inspecting the HasCustomFunction flag, the code determines whether SUMPRODUCT is available, applies it to the target cell only when supported, calculates the result, and saves the file.
// Keywords: Aspose.Cells SUMPRODUCT support | HasCustomFunction detection | conditional formula .NET | verify Excel function availability | runtime safe formula insertion
// Common Searches: Aspose.Cells how to test if SUMPRODUCT is supported | HasCustomFunction property example | prevent unsupported formula errors Aspose.Cells | check Excel function availability before use
// Developer Intent: Detect whether the SUMPRODUCT function exists in the current Aspose.Cells version and set the formula only if it is supported.
// Use Cases: Programmatically confirm support for a specific Excel function before writing it to a worksheet. | Avoid runtime exceptions caused by unsupported functions in older Aspose.Cells releases. | Implement fallback logic that switches to an alternative calculation when a function is missing.
// AI Prompts: Write C# code that checks for XLOOKUP support with Aspose.Cells and falls back to VLOOKUP if unavailable. | Create a utility that iterates over a list of formulas, uses HasCustomFunction to filter out unsupported ones, and applies the rest. | Generate a method that logs all custom or unsupported functions detected in a workbook using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, adds sample data, and uses a temporary cell to evaluate the SUMPRODUCT formula. By inspecting the HasCustomFunction flag, the code determines whether SUMPRODUCT is available, applies it to the target cell only when supported, calculates the result, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for the SUMPRODUCT calculation
        cells["A1"].PutValue(1);
        cells["A2"].PutValue(2);
        cells["B1"].PutValue(3);
        cells["B2"].PutValue(4);

        // Use a temporary cell to test whether SUMPRODUCT is supported
        Cell testCell = cells["C1"];
        testCell.Formula = "=SUMPRODUCT(A1:A2,B1:B2)";

        // If the formula contains an unsupported (custom) function, HasCustomFunction will be true
        bool sumProductSupported = !testCell.HasCustomFunction;

        if (sumProductSupported)
        {
            // SUMPRODUCT is supported – apply it to the actual target cell
            cells["D1"].Formula = "=SUMPRODUCT(A1:A2,B1:B2)";
            workbook.CalculateFormula(); // evaluate formulas
            Console.WriteLine("SUMPRODUCT result: " + cells["D1"].Value);
        }
        else
        {
            // Inform the user that the function is not available
            Console.WriteLine("SUMPRODUCT function is not supported in this Aspose.Cells version.");
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("SumProductCheck.xlsx");
    }
}
