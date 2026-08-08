// Title: Batch recalculate multiple Excel workbooks with a custom MYFUNC engine using Aspose.Cells for .NET
// Description: Iterates over a collection of Excel files, loads each workbook with Aspose.Cells, recalculates all formulas using a custom AbstractCalculationEngine that implements the MYFUNC function, and saves the updated workbooks to a separate folder while handling missing files and runtime errors.
// Keywords: Aspose.Cells | C# batch processing | custom calculation engine | MYFUNC function | CalculateFormula | multiple workbooks | Excel automation | AbstractCalculationEngine | recalculate formulas | save processed Excel files
// Common Searches: Aspose.Cells batch recalculate Excel files C# | custom function MYFUNC Aspose.Cells example | process multiple workbooks with custom engine .NET | how to use AbstractCalculationEngine in Aspose.Cells | save recalculated workbooks to a new folder
// Developer Intent: Load several Excel files, recalculate all formulas with a custom MYFUNC engine, and write the results to an output directory.
// Use Cases: Nightly batch job that updates financial models containing a proprietary MYFUNC across dozens of spreadsheets. | Automated preparation of monthly reports that rely on custom calculations before distribution to stakeholders. | Migration of a legacy spreadsheet library to a standardized format while preserving custom‑function results.
// AI Prompts: Generate C# code that adds type‑checking and logging for unsupported parameters in MyCustomEngine.Calculate. | Show how to extend MyCustomEngine to implement a new custom function SUMIFX that accepts a range and a criterion. | Create a unit test for MyCustomEngine.Calculate that verifies MYFUNC correctly sums values from both scalar arguments and cell ranges.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Iterates over a collection of Excel files, loads each workbook with Aspose.Cells, recalculates all formulas using a custom AbstractCalculationEngine that implements the MYFUNC function, and saves the updated workbooks to a separate folder while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        // Input workbook file paths
        var inputFiles = new List<string>
        {
            "Input1.xlsx",
            "Input2.xlsx"
        };

        // Output folder for processed workbooks
        string outputFolder = "Processed";
        Directory.CreateDirectory(outputFolder);

        // Custom calculation engine instance
        var customEngine = new MyCustomEngine();

        // Calculation options that use the custom engine
        var calcOptions = new CalculationOptions
        {
            CustomEngine = customEngine
        };

        // Process each workbook
        foreach (var inputPath in inputFiles)
        {
            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    continue;
                }

                // Load workbook
                Workbook wb = new Workbook(inputPath);

                // Recalculate all formulas using the custom engine
                wb.CalculateFormula(calcOptions);

                // Build output file name
                string fileName = Path.GetFileNameWithoutExtension(inputPath);
                string outputPath = Path.Combine(outputFolder, $"{fileName}_Processed.xlsx");

                // Save the processed workbook
                wb.Save(outputPath);
                Console.WriteLine($"Processed workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{inputPath}': {ex.Message}");
            }
        }
    }

    // Custom calculation engine implementing a sample function MYFUNC
    private class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle custom function named MYFUNC
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;

                // Iterate through all parameters of the function
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);

                    // Direct numeric value
                    if (param is double d)
                    {
                        sum += d;
                    }
                    // Parameter is a range (ReferredArea)
                    else if (param is ReferredArea area)
                    {
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            for (int c = area.StartColumn; c <= area.EndColumn; c++)
                            {
                                object val = area.GetValue(r, c);
                                if (val is double dv)
                                    sum += dv;
                            }
                        }
                    }
                }

                // Set the result of the custom function
                data.CalculatedValue = sum;
            }
        }
    }
}
