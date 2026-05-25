using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VerifyIfFunctionSupported
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Put test values in A1 and B1
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);

                // Use the IF function: if A1 > B1 then return "A>B" else return "A<=B"
                cells["C1"].Formula = "=IF(A1>B1,\"A>B\",\"A<=B\")";

                // Calculate formulas
                workbook.CalculateFormula();

                // Output the result to confirm IF is processed
                Console.WriteLine("Result of IF formula in C1: " + cells["C1"].StringValue);
                // Expected output: "A<=B" because 10 is not greater than 20
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