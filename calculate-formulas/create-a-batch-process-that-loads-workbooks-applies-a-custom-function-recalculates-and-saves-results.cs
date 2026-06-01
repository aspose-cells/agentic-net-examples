using System;
using System.IO;
using Aspose.Cells;

namespace BatchCustomFunctionDemo
{
    // Custom calculation engine that implements a user‑defined function named CUSTOMFUNC
    public class CustomCalcEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check for the custom function name (case‑insensitive)
            if (string.Equals(data.FunctionName, "CUSTOMFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Access the workbook that contains the formula via CalculationData.Workbook
                Workbook wb = data.Workbook;

                // Example logic: read value from cell A1 of the first worksheet,
                // write double of that value to cell C1, and return triple as the function result
                Worksheet ws = wb.Worksheets[0];
                double input = Convert.ToDouble(ws.Cells["A1"].Value ?? 0);
                ws.Cells["C1"].PutValue(input * 2);          // side‑effect
                data.CalculatedValue = input * 3;            // function return value
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder containing source workbooks
            string inputFolder = @"C:\InputWorkbooks";
            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Prepare calculation options with the custom engine
            CalculationOptions calcOptions = new CalculationOptions
            {
                CustomEngine = new CustomCalcEngine()
            };

            // Process each .xlsx file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Load workbook using the string constructor (lifecycle rule)
                Workbook wb = new Workbook(filePath);

                // Recalculate all formulas, invoking the custom function where used
                wb.CalculateFormula(calcOptions);

                // Build output file path (preserve original name)
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the updated workbook (lifecycle rule)
                wb.Save(outputPath);
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}