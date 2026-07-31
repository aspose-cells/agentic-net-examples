// Title: Define and Register a Custom MEDIAN Function in Aspose.Cells for .NET
// Description: Learn how to create a MedianEngine that extends AbstractCalculationEngine, mark the range argument as array‑mode with a custom function definition, register both with a Workbook, apply the formula =MEDIAN(A1:A6), calculate the result, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom function | C# median user‑defined function | register custom calculation engine | array‑mode parameter Aspose.Cells | calculate median range workbook | Aspose.Cells .NET example | custom statistical functions
// Common Searches: custom median function Aspose.Cells C# | how to register user defined function in Aspose.Cells | array mode custom function Aspose.Cells example | calculate median of a range with Aspose.Cells | Aspose.Cells custom calculation engine tutorial
// Developer Intent: Implement a user‑defined MEDIAN function that can be called in worksheet formulas and integrate it into the Aspose.Cells calculation pipeline.
// Use Cases: Compute the median of a column of numbers directly in a spreadsheet using =MEDIAN(range) after registering the custom function. | Add statistical or domain‑specific calculations (e.g., mode, percentile, custom aggregations) to Aspose.Cells without modifying the core library. | Automate report generation where custom formulas are required, evaluate them programmatically, and export the final workbook.
// AI Prompts: Generate C# code that defines a custom MEDIAN function for Aspose.Cells, registers it, and uses it in a worksheet formula. | Explain why array‑mode is needed for range parameters in Aspose.Cells custom functions and how to configure it. | Provide step‑by‑step instructions to test and debug a user‑defined median function in an Aspose.Cells .NET project.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomMedian
{
    // Custom calculation engine that implements the MEDIAN function
    // Learn how to create a MedianEngine that extends AbstractCalculationEngine, mark the range argument as array‑mode with a custom function definition, register both with a Workbook, apply the formula =MEDIAN(A1:A6), calculate the result, and save the workbook using Aspose.Cells for .NET.
    public class MedianEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Only handle the custom function named "MEDIAN"
            if (data.FunctionName.Equals("MEDIAN", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the first parameter (the range)
                object param = data.GetParamValue(0);

                // Collect numeric values from the parameter
                List<double> values = new List<double>();

                // If the parameter is already an array (array‑mode), iterate it
                if (param is object[,] array)
                {
                    foreach (object item in array)
                    {
                        if (item != null && double.TryParse(item.ToString(), out double d))
                            values.Add(d);
                    }
                }
                // Otherwise treat it as a single value
                else
                {
                    if (param != null && double.TryParse(param.ToString(), out double d))
                        values.Add(d);
                }

                // If there are no numeric values, result is 0
                if (values.Count == 0)
                {
                    data.CalculatedValue = 0.0;
                    return;
                }

                // Sort the values to compute median
                values.Sort();

                int n = values.Count;
                double median = (n % 2 == 1)
                    ? values[n / 2]
                    : (values[(n / 2) - 1] + values[n / 2]) / 2.0;

                data.CalculatedValue = median;
            }
        }
    }

    // Custom function definition that marks the first parameter of MEDIAN to be evaluated in array mode
    public class MedianFunctionDefinition : CustomFunctionDefinition
    {
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("MEDIAN", StringComparison.OrdinalIgnoreCase))
                return new[] { 0 };

            return base.GetArrayModeParameters(functionName);
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Populate sample data in A1:A6
                ws.Cells["A1"].PutValue(10);
                ws.Cells["A2"].PutValue(20);
                ws.Cells["A3"].PutValue(30);
                ws.Cells["A4"].PutValue(40);
                ws.Cells["A5"].PutValue(50);
                ws.Cells["A6"].PutValue(60);

                // Register the custom function definition so the parameter is processed in array mode
                wb.UpdateCustomFunctionDefinition(new MedianFunctionDefinition());

                // Set a formula that uses the custom MEDIAN function
                ws.Cells["B1"].Formula = "=MEDIAN(A1:A6)";

                // Prepare calculation options with the custom engine
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CustomEngine = new MedianEngine()
                };

                // Calculate all formulas in the workbook
                wb.CalculateFormula(calcOptions);

                // Output the result to console
                Console.WriteLine("Median of A1:A6 = " + ws.Cells["B1"].Value);

                // Save the workbook
                string outputPath = "MedianCustomFunctionDemo.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
