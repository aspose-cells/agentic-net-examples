using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCustomFunctionRangeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and get the first sheet
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // 2. Prepare a source range that the custom function will return
                // -------------------------------------------------
                // Fill A1:B2 with sample data
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["B1"].PutValue(30);
                cells["B2"].PutValue(40);

                // Create a Range object that refers to A1:B2
                AsposeRange sourceRange = cells.CreateRange("A1:B2");

                // -------------------------------------------------
                // 3. Define a cell that calls the custom function
                // -------------------------------------------------
                Cell formulaCell = cells["C1"];
                formulaCell.Formula = "=MYRANGE()";

                // -------------------------------------------------
                // 4. Set up calculation options with a custom engine
                // -------------------------------------------------
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new MyCustomEngine(sourceRange)
                };

                // -------------------------------------------------
                // 5. Calculate formulas – the custom engine will be invoked
                // -------------------------------------------------
                workbook.CalculateFormula(options);

                // -------------------------------------------------
                // 6. Retrieve the range returned by the custom function
                // -------------------------------------------------
                object result = formulaCell.Value; // should be a Range object

                if (result is AsposeRange returnedRange)
                {
                    // The Range.Value property holds a 2‑dimensional array of the cell values
                    object[,] values = (object[,])returnedRange.Value;

                    Console.WriteLine("Values returned by the custom function:");
                    for (int r = 0; r < values.GetLength(0); r++)
                    {
                        for (int c = 0; c < values.GetLength(1); c++)
                        {
                            Console.Write(values[r, c] + "\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    Console.WriteLine("The custom function did not return a Range.");
                }

                // -------------------------------------------------
                // 7. Save the workbook (optional, just to complete lifecycle)
                // -------------------------------------------------
                string outputPath = "CustomFunctionRangeResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // -----------------------------------------------------------------
        // Custom calculation engine that returns a predefined Range object
        // -----------------------------------------------------------------
        private class MyCustomEngine : AbstractCalculationEngine
        {
            private readonly AsposeRange _rangeToReturn;

            public MyCustomEngine(AsposeRange rangeToReturn)
            {
                _rangeToReturn = rangeToReturn;
            }

            public override void Calculate(CalculationData data)
            {
                // Check that the invoked function is the one we expect
                if (string.Equals(data.FunctionName, "MYRANGE", StringComparison.OrdinalIgnoreCase))
                {
                    // Return the predefined range as the calculated value
                    data.CalculatedValue = _rangeToReturn;
                }
                else
                {
                    // For any other function, let Aspose.Cells handle it (optional)
                    data.CalculatedValue = "#NAME?";
                }
            }
        }
    }
}