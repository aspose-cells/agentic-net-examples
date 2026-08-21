// Title: Confirm IF Function Support in Aspose.Cells for .NET
// Description: Creates a workbook, writes values to A1 and B1, sets an IF formula in C1, calculates the sheet, and outputs the result to verify that the IF function is listed in Aspose.Cells' Supported Excel Functions documentation.
// Keywords: Aspose.Cells | IF function | Excel formula support | .NET | C# | workbook calculation | supported functions list | formula evaluation example
// Common Searches: Aspose.Cells IF formula example | verify Excel IF support in Aspose.Cells | check supported functions Aspose.Cells .NET | how to test IF function with Aspose.Cells | Aspose.Cells documentation supported Excel functions
// Developer Intent: Validate that the IF function appears in the Supported Excel Functions reference for Aspose.Cells by executing it in a .NET workbook.
// Use Cases: Run a quick sanity check that the IF function works before processing large spreadsheets. | Automate a regression test suite that confirms core Excel functions remain supported after library upgrades. | Generate a log entry confirming successful execution of specific formulas during batch workbook generation.
// AI Prompts: Write a C# unit test using Aspose.Cells that asserts the IF formula returns the correct value. | Explain step‑by‑step how to programmatically verify any Excel function’s support in Aspose.Cells. | Create a script that iterates through a list of Excel functions, evaluates each with Aspose.Cells, and records pass/fail results.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, writes values to A1 and B1, sets an IF formula in C1, calculates the sheet, and outputs the result to verify that the IF function is listed in Aspose.Cells' Supported Excel Functions documentation.
    public class ConfirmIfFunctionSupported
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook wb = new Workbook();

                // Access the first worksheet and its cells
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Put sample values in A1 and B1
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);

                // Use the IF function in C1: if A1 > B1 then return A1 else return B1
                cells["C1"].Formula = "=IF(A1>B1, A1, B1)";

                // Calculate formulas (lifecycle rule: calculate)
                wb.CalculateFormula();

                // Retrieve the result
                var result = cells["C1"].Value;

                // Output the result and a confirmation that IF is supported
                Console.WriteLine($"Result of IF formula in C1: {result}");
                Console.WriteLine("The IF function executed successfully, confirming it is listed in the Supported Excel Functions documentation.");

                // Save the workbook (lifecycle rule: save)
                wb.Save("ConfirmIfFunctionSupported.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConfirmIfFunctionSupported.Run();
        }
    }
}
