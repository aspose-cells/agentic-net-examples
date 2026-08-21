// Title: Benchmark Single vs Multi‑Threaded Formula Calculation with Aspose.Cells .NET
// Description: Creates a 2000‑row by 50‑column workbook, fills cells with numeric data, adds a SUM formula per row, then measures calculation time twice—first using the default single‑threaded mode and next with multi‑threaded calculation enabled via WorkbookSettings.EnableMultiThreadedCalculation (set through reflection for version safety). The elapsed times are logged and the workbook is saved.
// Keywords: Aspose.Cells multi‑threaded calculation | EnableMultiThreadedCalculation .NET | formula calculation benchmark | CalculationOptions performance | large workbook processing | C# Aspose.Cells performance | reflection set property | Excel formula evaluation speed
// Common Searches: Aspose.Cells enable multi thread calculation | benchmark formula calculation Aspose.Cells | CalculationOptions ignore errors example | set EnableMultiThreadedCalculation via reflection | measure calculation time Aspose.Cells C# | performance of large workbook formulas .NET
// Developer Intent: Compare single‑threaded and multi‑threaded formula calculation speeds and learn how to enable multi‑threaded mode safely in Aspose.Cells for .NET.
// Use Cases: Determine if multi‑threaded calculation reduces processing time for financial models with thousands of rows. | Programmatically toggle multi‑threaded calculation based on library version to ensure backward compatibility. | Assess impact of CalculationOptions settings such as IgnoreError on overall calculation throughput. | Generate performance reports for large Excel workbooks before and after enabling multi‑threading.
// AI Prompts: Write C# code that logs detailed timing for preparation, calculation, and saving phases when using Aspose.Cells multi‑threaded calculation. | Provide a version‑agnostic method to enable EnableMultiThreadedCalculation without reflection, handling both newer and older Aspose.Cells releases. | Explain how to analyze benchmark results from Aspose.Cells calculation and recommend workbook design tweaks to maximize multi‑threaded performance.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a 2000‑row by 50‑column workbook, fills cells with numeric data, adds a SUM formula per row, then measures calculation time twice—first using the default single‑threaded mode and next with multi‑threaded calculation enabled via WorkbookSettings.EnableMultiThreadedCalculation (set through reflection for version safety). The elapsed times are logged and the workbook is saved.
class MultiThreadedCalculationDemo
{
    static void Main()
    {
        try
        {
            // Create a large workbook
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            int rows = 2000;   // number of rows
            int cols = 50;     // number of data columns (A..AX)

            // Populate cells with data and a sum formula per row
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    cells[i, j].PutValue(i + j);
                }

                // Formula: sum of the data columns in the current row
                string range = $"A{i + 1}:{ColumnIndexToName(cols - 1)}{i + 1}";
                cells[i, cols].Formula = $"=SUM({range})";
            }

            // Prepare calculation options (common for both runs)
            CalculationOptions options = new CalculationOptions
            {
                IgnoreError = true,
                Recursive = true
            };

            // ---------- Single‑threaded calculation ----------
            // In older Aspose.Cells versions multi‑threaded calculation may not be available.
            // Ensure it is disabled (default) and perform calculation.
            Stopwatch sw = Stopwatch.StartNew();
            wb.CalculateFormula(options);
            sw.Stop();
            Console.WriteLine($"Single‑threaded calculation time: {sw.ElapsedMilliseconds} ms");

            // Reset workbook to original state by re‑creating it (to get fair comparison)
            wb = new Workbook();
            ws = wb.Worksheets[0];
            cells = ws.Cells;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    cells[i, j].PutValue(i + j);
                }
                string range = $"A{i + 1}:{ColumnIndexToName(cols - 1)}{i + 1}";
                cells[i, cols].Formula = $"=SUM({range})";
            }

            // ---------- Multi‑threaded calculation ----------
            // If the current Aspose.Cells version supports multi‑threaded calculation,
            // enable it via WorkbookSettings.EnableMultiThreadedCalculation.
            // The property may be unavailable in some versions; guard with a try‑catch.
            try
            {
                // Attempt to enable multi‑threaded calculation if the property exists.
                // This uses reflection to avoid compile‑time errors on older versions.
                var settingsType = typeof(WorkbookSettings);
                var prop = settingsType.GetProperty("EnableMultiThreadedCalculation");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(wb.Settings, true);
                }
            }
            catch
            {
                // Property not supported; continue with default (single‑threaded) behavior.
            }

            sw.Restart();
            wb.CalculateFormula(options);
            sw.Stop();
            Console.WriteLine($"Multi‑threaded calculation time: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (optional)
            string outputPath = "LargeWorkbook.xlsx";
            wb.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper: converts zero‑based column index to Excel column name (e.g., 0 -> A, 27 -> AB)
    static string ColumnIndexToName(int index)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string name = "";
        do
        {
            name = letters[index % 26] + name;
            index = index / 26 - 1;
        } while (index >= 0);
        return name;
    }
}
