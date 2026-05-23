using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate a large dataset with formulas (example: 2000 rows)
            int totalRows = 2000;
            for (int i = 0; i < totalRows; i++)
            {
                // Simple numeric value in column A
                cells[i, 0].PutValue(i + 1);

                // Formula in column B that depends on column A
                cells[i, 1].Formula = $"=A{i + 1}*2";
            }

            // Fast formula calculation flag may not be available in older versions;
            // if present, it can be enabled as shown below (commented out for compatibility)
            // workbook.Settings.EnableFastFormulaCalculation = true;

            // Prepare calculation options (default options are sufficient)
            CalculationOptions calcOptions = new CalculationOptions();

            // Calculate each formula cell individually
            for (int i = 0; i < totalRows; i++)
            {
                cells[i, 1].Calculate(calcOptions);
            }

            // Determine output file path and ensure the directory exists
            string outputFile = Path.Combine(Environment.CurrentDirectory, "FastFormulaCalculationResult.xlsx");
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with calculated results
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {outputFile}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}