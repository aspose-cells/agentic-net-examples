// Title: Extract Array Formula Text with FORMULATEXT in Aspose.Cells for .NET
// Description: This C# example creates a workbook, fills A1:B3 with numbers, applies the array formula `SUM(A1:A3*B1:B3)` to C1 using SetArrayFormula, then uses CalculateFormula("=FORMULATEXT(C1)") to obtain the exact formula string for debugging, prints it, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | array formula | FORMULATEXT | SetArrayFormula | CalculateFormula | extract formula text | Excel formula debugging | retrieve formula string
// Common Searches: Aspose.Cells get formula text C# | How to use FORMULATEXT with Aspose.Cells | Extract array formula as string in .NET | Debug Aspose.Cells array formulas | Retrieve Excel formula programmatically
// Developer Intent: Obtain the literal text of an array formula applied to a cell to verify or debug its implementation.
// Use Cases: Log the exact formula of a cell to confirm that SetArrayFormula was applied correctly. | Compare the extracted formula against an expected pattern in automated tests. | Generate documentation that includes the original Excel formulas without opening the file.
// AI Prompts: Show C# code that extracts the FORMULATEXT of a cell containing an array formula using Aspose.Cells. | Provide a .NET example for retrieving and printing the text of a complex array formula in an Excel workbook. | Explain how to debug Aspose.Cells array formulas by obtaining their textual representation.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, fills A1:B3 with numbers, applies the array formula `SUM(A1:A3*B1:B3)` to C1 using SetArrayFormula, then uses CalculateFormula("=FORMULATEXT(C1)") to obtain the exact formula string for debugging, prints it, and saves the workbook.
    public class ExtractArrayFormulaTextDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data that the array formula will use
                sheet.Cells["A1"].PutValue(1);
                sheet.Cells["A2"].PutValue(2);
                sheet.Cells["A3"].PutValue(3);
                sheet.Cells["B1"].PutValue(4);
                sheet.Cells["B2"].PutValue(5);
                sheet.Cells["B3"].PutValue(6);

                // Define a complex array formula
                string arrayFormula = "SUM(A1:A3*B1:B3)";

                // Apply the array formula to cell C1
                sheet.Cells["C1"].SetArrayFormula(arrayFormula, 0, 0);

                // Retrieve the textual representation of the array formula
                object formulaText = sheet.CalculateFormula("=FORMULATEXT(C1)");

                Console.WriteLine("Extracted array formula text: " + (formulaText?.ToString() ?? "null"));

                // Save the workbook
                string outputPath = "ExtractArrayFormulaTextDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Program entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ExtractArrayFormulaTextDemo.Run();
        }
    }
}
