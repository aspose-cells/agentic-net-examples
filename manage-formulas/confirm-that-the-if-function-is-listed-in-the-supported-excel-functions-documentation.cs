// Title: Verify IF Function Support in Aspose.Cells for .NET
// Description: C# example that creates a workbook, writes values to A1 and B1, sets an IF formula in C1, runs CalculateFormula, prints the result, and optionally saves the file to confirm that the IF function is processed correctly by Aspose.Cells.
// Keywords: Aspose.Cells | IF function | .NET | C# formula calculation | supported Excel functions | Excel IF example | Aspose.Cells documentation | formula evaluation
// Common Searches: Aspose.Cells IF function support | does Aspose.Cells evaluate IF formulas | list of supported Excel functions Aspose.Cells | C# verify Excel IF in Aspose.Cells | Aspose.Cells calculate IF formula
// Developer Intent: Confirm that the IF function is recognized and correctly evaluated by Aspose.Cells for .NET.
// Use Cases: Programmatically ensure critical logical formulas work after migration. | Automated testing of formula support during CI builds. | Validate end‑to‑end workbook processing when using conditional logic.
// AI Prompts: Write an NUnit test in C# that inserts an IF formula, calculates it with Aspose.Cells, and asserts the expected output. | Generate a PowerShell script that checks the Aspose.Cells documentation for the IF function entry and fails if missing. | Explain how to retrieve the catalog of supported Excel functions from Aspose.Cells via its API or documentation.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, writes values to A1 and B1, sets an IF formula in C1, runs CalculateFormula, prints the result, and optionally saves the file to confirm that the IF function is processed correctly by Aspose.Cells.
    public class VerifyIfFunctionSupported
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook wb = new Workbook();

                // Access the first worksheet and its cells
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Put test values in cells A1 and B1
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);

                // Use the IF function in cell C1
                // IF(A1 > B1, "A is greater", "B is greater or equal")
                cells["C1"].Formula = "=IF(A1>B1,\"A is greater\",\"B is greater or equal\")";

                // Calculate formulas (supported functions are processed here)
                wb.CalculateFormula();

                // Output the result to confirm that IF was evaluated correctly
                Console.WriteLine("Result of IF function in C1: " + cells["C1"].StringValue);
                // Expected output: "B is greater or equal" because 10 is not greater than 20

                // Save the workbook (lifecycle: save) – optional, just to complete the lifecycle steps
                wb.Save("IfFunctionSupported.xlsx", SaveFormat.Xlsx);
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
            VerifyIfFunctionSupported.Run();
        }
    }
}
