// Title: C# memory profiling of Aspose.Cells default vs custom calculation engine during large workbook formula evaluation
// Common Searches: C# how to measure memory usage of Aspose.Cells CalculateFormula for large Excel files | compare memory consumption of default and custom calculation engines in Aspose.Cells .NET | profile heap allocation when evaluating formulas with Aspose.Cells workbook | benchmark Aspose.Cells formula calculation time and memory in .NET | Aspose.Cells .NET custom calculation engine memory impact assessment
// Tags: Aspose.Cells default calculation engine memory profiling | custom calculation engine heap usage Aspose.Cells | large workbook formula evaluation performance .NET | GC.GetTotalMemory Aspose.Cells benchmark | Workbook.CalculateFormula memory consumption

using System;
using System.Diagnostics;
using Aspose.Cells;

// The sample creates a 10,000‑row by 50‑column workbook, fills it with numeric data, adds a SUM formula per row, and measures heap usage and elapsed time before and after calling Workbook.CalculateFormula() using the default engine. It then clones the workbook, optionally assigns a custom calculation engine, repeats the memory and timing measurements, and prints both results, allowing a direct comparison of memory consumption and performance between the default and a custom calculation engine.
class MemoryConsumptionComparison
{
    static void Main()
    {
        try
        {
            // ---------- Create a large workbook ----------
            Workbook workbook = new Workbook();                     // {CreateWorkbook}
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the sheet with a large amount of data
            const int rows = 10000;
            const int cols = 50;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    sheet.Cells[i, j].PutValue(i + j);
                }
            }

            // Add formulas that will be evaluated
            for (int i = 0; i < rows; i++)
            {
                // Example formula: sum of the row
                sheet.Cells[i, cols].Formula = $"=SUM(A{i + 1}:AX{i + 1})";
            }

            // ---------- Profile default calculation engine ----------
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memBeforeDefault = GC.GetTotalMemory(true);
            Stopwatch swDefault = Stopwatch.StartNew();

            workbook.CalculateFormula();                                   // default engine evaluation

            swDefault.Stop();
            long memAfterDefault = GC.GetTotalMemory(true);
            long memUsedDefault = memAfterDefault - memBeforeDefault;

            Console.WriteLine($"Default Engine -> Time: {swDefault.ElapsedMilliseconds} ms, Memory: {memUsedDefault / 1024} KB");

            // ---------- Profile cloned workbook (identical data) ----------
            Workbook customWorkbook = new Workbook();
            customWorkbook.Copy(workbook);                          // {CopyWorkbook}

            // Placeholder for custom calculation engine assignment, if needed
            // customWorkbook.Settings.CalcEngine = new YourCustomEngine();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memBeforeCustom = GC.GetTotalMemory(true);
            Stopwatch swCustom = Stopwatch.StartNew();

            customWorkbook.CalculateFormula();                             // evaluation

            swCustom.Stop();
            long memAfterCustom = GC.GetTotalMemory(true);
            long memUsedCustom = memAfterCustom - memBeforeCustom;

            Console.WriteLine($"Custom Engine  -> Time: {swCustom.ElapsedMilliseconds} ms, Memory: {memUsedCustom / 1024} KB");

            // ---------- Save workbooks (optional) ----------
            // workbook.Save("DefaultEngine.xlsx");                 // {SaveWorkbook}
            // customWorkbook.Save("CustomEngine.xlsx");            // {SaveWorkbook}
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
