// Title: Measure Aspose.Cells Workbook.CalculateFormula performance with EnableThreadedCalculation on and off in C#
// Common Searches: how to measure Aspose.Cells CalculateFormula speed with threaded calculation enabled in C# | benchmark Workbook.CalculateFormula execution time using EnableThreadedCalculation true versus false | C# code to compare multithreaded and single‑threaded formula calculation in Aspose.Cells | Aspose.Cells performance test for EnableThreadedCalculation setting | using Stopwatch to profile Aspose.Cells formula calculation in .NET
// Tags: Aspose.Cells CalculateFormula performance benchmarking | EnableThreadedCalculation comparison in .NET | C# Stopwatch profiling of workbook calculation | multithreaded vs single‑threaded formula evaluation Aspose.Cells | WorkbookSettings reflection property toggle

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example creates or loads an input.xlsx workbook, toggles WorkbookSettings.EnableThreadedCalculation to true and false via reflection, runs Workbook.CalculateFormula for each configuration, measures the elapsed time with Stopwatch, and outputs the threaded and non‑threaded calculation durations to help determine the optimal setting.
class Program
{
    static void Main()
    {
        // Path to the workbook to be used for the test
        string workbookPath = "input.xlsx";

        // Ensure the input file exists; create a simple workbook if it does not.
        if (!File.Exists(workbookPath))
        {
            try
            {
                var wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue(1);
                // Use the Formula property (compatible with all versions)
                wb.Worksheets[0].Cells["A2"].Formula = "=A1*2";
                wb.Save(workbookPath);
                Console.WriteLine($"Created placeholder workbook at '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create placeholder workbook: {ex.Message}");
                return;
            }
        }

        // ------------------------------
        // Test with EnableThreadedCalculation = true (if supported)
        // ------------------------------
        try
        {
            // Load the workbook
            Workbook threadedWb = new Workbook(workbookPath);

            // Attempt to enable threaded calculation via reflection (compatible with older versions)
            try
            {
                var prop = typeof(WorkbookSettings).GetProperty("EnableThreadedCalculation");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(threadedWb.Settings, true);
                }
            }
            catch (Exception reflEx)
            {
                Console.WriteLine($"Unable to set EnableThreadedCalculation via reflection: {reflEx.Message}");
            }

            // Measure calculation time
            Stopwatch swThreaded = Stopwatch.StartNew();
            threadedWb.CalculateFormula(); // Perform calculation
            swThreaded.Stop();

            Console.WriteLine($"Threaded calculation time: {swThreaded.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during threaded calculation: {ex.Message}");
        }

        // ------------------------------
        // Test with EnableThreadedCalculation = false (if supported)
        // ------------------------------
        try
        {
            // Load the workbook again to avoid cached results
            Workbook nonThreadedWb = new Workbook(workbookPath);

            // Attempt to disable threaded calculation via reflection
            try
            {
                var prop = typeof(WorkbookSettings).GetProperty("EnableThreadedCalculation");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(nonThreadedWb.Settings, false);
                }
            }
            catch (Exception reflEx)
            {
                Console.WriteLine($"Unable to set EnableThreadedCalculation via reflection: {reflEx.Message}");
            }

            // Measure calculation time
            Stopwatch swNonThreaded = Stopwatch.StartNew();
            nonThreadedWb.CalculateFormula(); // Perform calculation
            swNonThreaded.Stop();

            Console.WriteLine($"Non-threaded calculation time: {swNonThreaded.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during non-threaded calculation: {ex.Message}");
        }
    }
}
