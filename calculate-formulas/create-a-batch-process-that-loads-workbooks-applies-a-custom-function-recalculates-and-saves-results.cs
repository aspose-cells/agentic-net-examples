// Title: Batch recalculate Excel workbooks with a custom DOUBLE(x) function using Aspose.Cells for .NET
// Description: A C# utility that scans a source directory, loads each .xls or .xlsx file with Aspose.Cells, registers a custom calculation engine implementing DOUBLE(x) = x × 2, recalculates all formulas, and saves the updated workbooks to a target folder while preserving original file names.
// Keywords: Aspose.Cells batch processing | C# custom calculation engine | Excel workbook recalculate | DOUBLE custom function | load and save Excel files .NET | process multiple workbooks programmatically | calculate formulas with custom engine | directory based Excel automation
// Common Searches: Aspose.Cells custom function example C# | Batch recalculate formulas in Excel files using Aspose.Cells | How to add a DOUBLE(x) function to Aspose.Cells calculation engine | Save processed workbooks to a different folder Aspose.Cells | C# loop through Excel files and recalculate formulas
// Developer Intent: Create a C# batch routine that loads each workbook, applies a custom DOUBLE function during formula calculation, and writes the refreshed files to an output directory.
// Use Cases: Nightly re‑calculation of financial models where a custom scaling factor (DOUBLE) must be applied before archiving. | Automated processing of daily sales reports that contain the DOUBLE function, updating values and depositing the results in a reporting folder. | Bulk adjustment of engineering spreadsheets in a shared drive, applying a custom multiplier and saving the modified files for downstream analysis.
// AI Prompts: Generate C# code that adds a SUMSQ(x,y) custom function to the existing DoubleFunctionEngine and updates the batch loop to use it. | Explain how Aspose.Cells invokes CustomEngine during Workbook.CalculateFormula and suggest debugging steps for parameter conversion issues. | Write NUnit tests for DoubleFunctionEngine verifying that DOUBLE returns twice the input for integers, doubles, and numeric strings.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchCustomFunctionDemo
{
    // Custom calculation engine that implements a simple function DOUBLE(x) = x * 2
    // A C# utility that scans a source directory, loads each .xls or .xlsx file with Aspose.Cells, registers a custom calculation engine implementing DOUBLE(x) = x × 2, recalculates all formulas, and saves the updated workbooks to a target folder while preserving original file names.
    public class DoubleFunctionEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is our custom function
            if (data.FunctionName.Equals("DOUBLE", StringComparison.OrdinalIgnoreCase))
            {
                // Get the first parameter value
                object param = data.GetParamValue(0);
                double value = Convert.ToDouble(param);
                // Set the calculated result
                data.CalculatedValue = value * 2;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder containing source workbooks
            string sourceFolder = @"C:\InputWorkbooks";
            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (xlsx and xls)
            string[] files = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in files)
            {
                // Load workbook using the string constructor (load rule)
                Workbook workbook = new Workbook(filePath);

                // Prepare calculation options with our custom engine
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CustomEngine = new DoubleFunctionEngine()
                };

                // Recalculate all formulas (including custom functions) in the workbook
                workbook.CalculateFormula(calcOptions);

                // Build output file path (preserve original name)
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the processed workbook using the Save(string) method (save rule)
                workbook.Save(outputPath);
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
