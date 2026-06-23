using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace CustomEngineDemo
{
    // Custom calculation engine that returns a Range object for the function GETRANGE()
    public class MyRangeEngine : AbstractCalculationEngine
    {
        // Called for each custom function encountered during formula evaluation
        public override void Calculate(CalculationData data)
        {
            // Verify the function name (case‑insensitive)
            if (data.FunctionName.Equals("GETRANGE", StringComparison.OrdinalIgnoreCase))
            {
                // Worksheet where the function is evaluated
                Worksheet ws = data.Worksheet;

                // Create a Range that refers to cells B1:C2 on the same worksheet
                AsposeRange resultRange = ws.Cells.CreateRange("B1:C2");

                // Set the calculated value of the function to the Range object
                data.CalculatedValue = resultRange;
            }
        }

        // Ensure the function is recalculated for each call (volatile function)
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("GETRANGE", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (optional, just for illustration)
                sheet.Cells["A1"].PutValue(1);
                sheet.Cells["A2"].PutValue(2);
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["C1"].PutValue(30);
                sheet.Cells["C2"].PutValue(40);

                // Set a formula that invokes the custom function GETRANGE()
                sheet.Cells["D1"].Formula = "=GETRANGE()";

                // Configure calculation options to use the custom engine
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new MyRangeEngine()
                };

                // Perform calculation
                workbook.CalculateFormula(options);

                // Retrieve the result from D1; it should be a Range object
                object rawResult = sheet.Cells["D1"].Value;
                AsposeRange returnedRange = rawResult as AsposeRange;

                if (returnedRange != null)
                {
                    // Calculate start/end indices using Range properties
                    int startRow = returnedRange.FirstRow;
                    int startColumn = returnedRange.FirstColumn;
                    int endRow = startRow + returnedRange.RowCount - 1;
                    int endColumn = startColumn + returnedRange.ColumnCount - 1;

                    Console.WriteLine("Custom function returned range:");
                    Console.WriteLine($"Start: Row {startRow + 1}, Column {startColumn + 1}");
                    Console.WriteLine($"End:   Row {endRow + 1}, Column {endColumn + 1}");
                }
                else
                {
                    Console.WriteLine("The custom function did not return a Range.");
                }

                // Save the workbook (ensure the directory exists)
                string outputPath = "RangeCustomFunctionDemo.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}