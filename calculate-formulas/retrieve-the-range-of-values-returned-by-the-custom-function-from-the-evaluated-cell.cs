// Title: Extract values from a Range returned by a custom function in Aspose.Cells (.NET)
// Description: Demonstrates creating a workbook, defining a custom function (MYRANGE) that returns the range A1:B2, calculating formulas with a custom engine, and retrieving the evaluated cell's Value. The result is cast to Aspose.Cells.Range, allowing access to its 2‑dimensional value array or iteration over individual cells.
// Keywords: Aspose.Cells custom function | C# return range from formula | AsposeRange value extraction | calculate formulas Aspose.Cells | .NET workbook custom engine | retrieve 2D array from cell | Aspose.Cells range object | Excel custom function C# | Aspose.Cells API example
// Common Searches: Aspose.Cells get values from custom function range | C# read AsposeRange returned by MYRANGE | How to cast formula cell value to Range in Aspose.Cells | Extract 2D object array from Aspose.Cells formula | Iterate cells of a Range returned by custom engine
// Developer Intent: The developer needs to obtain each cell’s value from the Range object that a custom function returns after formula evaluation.
// Use Cases: After workbook.CalculateFormula, cast the formula cell’s Value to AsposeRange and read its Value property as an object[,] for bulk processing. | When the Range’s Value is not a 2‑D array, enumerate the cells in the returned Range to access individual Name and Value pairs. | Save the workbook after extracting range data for reporting, further calculations, or exporting to other formats.
// AI Prompts: Write C# code that converts a Range object returned by a custom Aspose.Cells function into a two‑dimensional object array. | Show how to handle non‑contiguous ranges returned by a custom calculation engine and collect their values into a list. | Create a helper method that transforms an AsposeRange into List<List<object>> for easy consumption in .NET applications.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCustomFunctionDemo
{
    // Demonstrates creating a workbook, defining a custom function (MYRANGE) that returns the range A1:B2, calculating formulas with a custom engine, and retrieving the evaluated cell's Value. The result is cast to Aspose.Cells.Range, allowing access to its 2‑dimensional value array or iteration over individual cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Populate some sample data that will be referenced by the custom function
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["B1"].PutValue(30);
                cells["B2"].PutValue(40);

                // 3. Set a formula that calls the custom function MYRANGE()
                //    The function will return the range A1:B2 as its calculated value
                Cell formulaCell = cells["C1"];
                formulaCell.Formula = "=MYRANGE()";

                // 4. Configure calculation options to use a custom engine
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new MyRangeEngine()
                };

                // 5. Calculate all formulas in the workbook
                workbook.CalculateFormula(options);

                // 6. Retrieve the value from the evaluated cell.
                //    Since the custom function returns a Range object, the cell's Value will be that Range.
                object result = formulaCell.Value;

                if (result is AsposeRange returnedRange)
                {
                    // 7. The returned range may contain multiple cells.
                    //    Its Value property holds a 2‑dimensional array with the cell values.
                    Console.WriteLine("Custom function returned a range with the following values:");
                    if (returnedRange.Value is object[,] values)
                    {
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
                        // Fallback: iterate through cells if Value is not a 2‑D array
                        foreach (Cell cell in returnedRange)
                        {
                            Console.WriteLine($"{cell.Name}: {cell.Value}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Unexpected result type: " + (result?.GetType().FullName ?? "null"));
                }

                // 8. Save the workbook (optional, just to demonstrate lifecycle compliance)
                string outputPath = "CustomFunctionResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Custom calculation engine that handles the MYRANGE function
        private class MyRangeEngine : AbstractCalculationEngine
        {
            public override void Calculate(CalculationData data)
            {
                // Verify that we are processing the expected custom function
                if (string.Equals(data.FunctionName, "MYRANGE", StringComparison.OrdinalIgnoreCase))
                {
                    // Obtain the worksheet where the function is evaluated
                    Worksheet ws = data.Worksheet;

                    // Create a range that we want to return (A1:B2 in this example)
                    AsposeRange rangeToReturn = ws.Cells.CreateRange("A1:B2");

                    // Set the calculated value of the function to the Range object
                    data.CalculatedValue = rangeToReturn;
                }
            }
        }
    }
}
