// Title: Interrupt Aspose.Cells workbook calculation in C# with the BeforeCalculate event when a cell exceeds a threshold
// Common Searches: how to stop Aspose.Cells calculation when cell B2 is greater than 500 in C# | using BeforeCalculate event to interrupt formula evaluation in Aspose.Cells .NET | conditional workbook calculation interruption example with Aspose.Cells | prevent CalculateFormula from running based on a numeric cell condition in C# | Aspose.Cells C# example to halt calculation if a cell meets a specific criteria
// Tags: Aspose.Cells BeforeCalculate interrupt | conditional workbook calculation stop | cell value threshold interrupt | C# Aspose.Cells calculation control | interrupt formula evaluation based on cell value

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates an Excel workbook, ensures cell B2 exists, assigns a numeric value, checks if the value exceeds 500, and uses the BeforeCalculate event to call Interrupt() to halt the workbook.CalculateFormula() call, otherwise it proceeds with calculation and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the target cell exists (creates it if necessary)
                Cell targetCell = sheet.Cells["B2"];

                // Example: set a value for demonstration purposes
                // Remove or modify this line as needed
                targetCell.PutValue(600);

                // Check the condition before performing calculation
                bool interrupt = false;
                if (targetCell.Value != null &&
                    targetCell.Type == CellValueType.IsNumeric &&
                    targetCell.IntValue > 500)
                {
                    interrupt = true;
                }

                if (!interrupt)
                {
                    // Perform full workbook calculation
                    workbook.CalculateFormula();
                }
                else
                {
                    Console.WriteLine("Calculation interrupted because B2 > 500.");
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
