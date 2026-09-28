// Title: Measure the effect of Aspose.Cells calculation chain on formula evaluation speed across multiple worksheets in C#
// AI Prompts: Generate C# code with Aspose.Cells that loops through every worksheet, disables the calculation chain, times Workbook.CalculateFormula, then enables the chain and times it again, printing both durations. | Extend the sample to record the elapsed time for each worksheet separately, store the results in a DataTable, and write the performance data to a CSV file. | Create a version of the program that accepts the row count per sheet as a command‑line argument and compares calculation‑chain performance for small (100 rows), medium (1,000 rows), and large (10,000 rows) workbooks.
// Common Searches: how to benchmark Aspose.Cells formula calculation speed with calculation chain enabled | C# Aspose.Cells measure CalculateFormula execution time per worksheet | compare performance of fast formula calculation versus normal in Aspose.Cells .NET | disable and enable calculation chain in Aspose.Cells to improve recalc time
// Tags: Aspose.Cells calculation chain performance | benchmark Workbook.CalculateFormula .NET | measure formula evaluation time Aspose.Cells | iterate worksheets fast calculation Aspose.Cells | export performance data to CSV C#

using System;
using System.Diagnostics;
using Aspose.Cells;

// The program creates a workbook with two worksheets, fills each with a chain of dependent formulas, disables the calculation chain, measures the time required for Workbook.CalculateFormula, then enables the calculation chain, measures again, compares the two timings, outputs the results, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add an extra worksheet to have multiple sheets
        workbook.Worksheets.Add();

        // Populate each worksheet with sample data and formulas
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Simple chain of formulas to create dependency
            cells["A1"].PutValue(1);
            for (int i = 2; i <= 1000; i++)
            {
                cells[$"A{i}"].Formula = $"=A{i - 1}+1";
            }

            // A summary formula that sums the column
            cells["B1"].Formula = "=SUM(A1:A1000)";
        }

        // Ensure fast calculation (calculation chain) is disabled initially
        workbook.Settings.FormulaSettings.EnableCalculationChain = false;

        // Measure calculation time without the calculation chain
        Stopwatch sw = Stopwatch.StartNew();
        workbook.CalculateFormula();
        sw.Stop();
        long timeWithoutChain = sw.ElapsedMilliseconds;
        Console.WriteLine($"Calculation time without chain: {timeWithoutChain} ms");

        // Enable fast formula calculation (calculation chain)
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Measure calculation time with the calculation chain enabled
        sw.Restart();
        workbook.CalculateFormula();
        sw.Stop();
        long timeWithChain = sw.ElapsedMilliseconds;
        Console.WriteLine($"Calculation time with chain: {timeWithChain} ms");

        // Compare the two timings
        if (timeWithChain < timeWithoutChain)
            Console.WriteLine("Fast calculation improved performance.");
        else
            Console.WriteLine("Fast calculation did not improve performance.");

        // Save the workbook (optional)
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}
