using System;
using System.Diagnostics;
using Aspose.Cells;

public class CalculationChainDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate a large number of formulas to see performance impact
        int totalRows = 5000;
        cells["A1"].PutValue(1);
        for (int i = 2; i <= totalRows; i++)
        {
            // Incremental values in column A
            cells[$"A{i}"].PutValue(i);

            // Each cell in column B sums the range A1:A{i}
            cells[$"B{i}"].Formula = $"=SUM(A1:A{i})";
        }

        // Helper method to calculate the worksheet and measure elapsed time
        void CalculateAndMeasure(bool enableChain)
        {
            // Enable or disable the calculation chain
            workbook.Settings.FormulaSettings.EnableCalculationChain = enableChain;

            // Start timing
            Stopwatch sw = Stopwatch.StartNew();

            // Recalculate only the current worksheet
            worksheet.CalculateFormula(new CalculationOptions(), true);

            // Stop timing
            sw.Stop();

            Console.WriteLine($"EnableCalculationChain = {enableChain}: {sw.ElapsedMilliseconds} ms");
        }

        // Measure with calculation chain enabled
        CalculateAndMeasure(true);

        // Measure with calculation chain disabled
        CalculateAndMeasure(false);
    }
}