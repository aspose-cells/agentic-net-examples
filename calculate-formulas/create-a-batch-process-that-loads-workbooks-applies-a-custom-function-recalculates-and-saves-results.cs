// Title: Batch recalculate formulas with a custom MYFUNC engine across multiple Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over a list of .xlsx file paths, loads each workbook with Aspose.Cells, assigns a custom calculation engine that implements a MYFUNC function, recalculates all formulas, and saves the workbook with a "_Processed" suffix. | Create a custom Aspose.Cells calculation engine in C# that intercepts the MYFUNC function, reads the value of cell A1 from the first worksheet, returns its square, and integrate this engine into a batch loop that processes several Excel files.
// Common Searches: how to apply a custom calculation engine to recalculate formulas in several Excel files with Aspose.Cells C# | batch processing Excel workbooks with user‑defined functions using Aspose.Cells .NET | C# example for loading multiple .xlsx files, running custom MYFUNC, and saving processed copies | Aspose.Cells calculate all formulas with custom function for each workbook in a folder
// Tags: batch calculate formulas Aspose.Cells C# | custom calculation engine MYFUNC Aspose.Cells | load multiple .xlsx files Aspose.Cells | save processed workbook with suffix C# | recalculate all formulas with custom function Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample program loops through a collection of Excel files, loads each workbook with Aspose.Cells, attaches a custom calculation engine that implements a MYFUNC function (squaring the value from cell A1), recalculates every formula, and saves the updated workbook with a "_Processed" suffix, while handling missing files and runtime errors.
class BatchProcessor
{
    static void Main()
    {
        // List of workbook files to process
        var inputFiles = new List<string>
        {
            "Input1.xlsx",
            "Input2.xlsx"
            // Add more file paths as needed
        };

        foreach (var inputPath in inputFiles)
        {
            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {inputPath}. Skipping.");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Set up calculation options with a custom engine
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new MyCustomEngine()
                };

                // Recalculate all formulas (including custom functions)
                workbook.CalculateFormula(options);

                // Build output file name (adds "_Processed" suffix)
                string directory = Path.GetDirectoryName(inputPath) ?? string.Empty;
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                string outputPath = Path.Combine(directory, $"{fileNameWithoutExt}_Processed.xlsx");

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{inputPath}': {ex.Message}");
            }
        }
    }

    // Custom calculation engine implementing a sample function named MYFUNC
    private class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check for the custom function name (case‑insensitive)
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Access the workbook where the function is being evaluated
                Workbook wb = data.Workbook;

                // Example logic: read a value from cell A1 of the first worksheet
                double input = Convert.ToDouble(wb.Worksheets[0].Cells["A1"].Value);

                // Perform custom calculation (square the input)
                data.CalculatedValue = input * input;
            }
        }
    }
}
