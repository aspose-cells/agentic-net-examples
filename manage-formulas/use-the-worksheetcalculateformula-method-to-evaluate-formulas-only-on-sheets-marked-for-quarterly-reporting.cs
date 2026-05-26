using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsQuarterlyReport
{
    public class QuarterlyFormulaEvaluator
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "InputReport.xlsx";
            const string outputPath = "OutputReport.xlsx";

            try
            {
                // Ensure the input file exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    bool isQuarterly = false;

                    // Determine quarterly sheets by name pattern
                    if (sheet.Name.IndexOf("Quarterly", StringComparison.OrdinalIgnoreCase) >= 0)
                        isQuarterly = true;

                    // Uncomment the block below if using a custom property instead of naming convention
                    /*
                    if (sheet.CustomProperties.Contains("ReportType"))
                    {
                        string reportType = sheet.CustomProperties["ReportType"]?.ToString();
                        if (string.Equals(reportType, "Quarterly", StringComparison.OrdinalIgnoreCase))
                            isQuarterly = true;
                    }
                    */

                    if (isQuarterly)
                    {
                        // Calculate formulas in the worksheet with custom options
                        CalculationOptions opts = new CalculationOptions { Recursive = true };
                        sheet.CalculateFormula(opts, true);
                    }
                }

                // Save the workbook after calculations
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}