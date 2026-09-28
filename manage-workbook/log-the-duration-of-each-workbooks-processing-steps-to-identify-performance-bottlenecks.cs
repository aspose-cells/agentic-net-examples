// Title: Measure and log execution time of loading, formula calculation, AutoFit columns, and saving an Excel workbook with Aspose.Cells in C#
// AI Prompts: Wrap each Aspose.Cells workbook operation (new Workbook, CalculateFormula, AutoFitColumns, Save) with a Stopwatch and output the elapsed milliseconds. | Implement a helper method that accepts an Action representing a workbook step, returns the elapsed time, and use it to profile all processing stages. | Aggregate the timings of each step into a collection and print a formatted performance summary after the workbook is saved.
// Common Searches: how to profile Aspose.Cells workbook processing time in C# | C# code to measure load and save duration of an Excel file using Aspose.Cells | benchmarking formula calculation speed with Aspose.Cells .NET | record step-by-step execution times for AutoFitColumns in Aspose.Cells | total processing time for Excel workbook using Aspose.Cells and Stopwatch
// Tags: Aspose.Cells workbook load timing | Aspose.Cells formula calculation benchmark | Aspose.Cells AutoFitColumns performance | Aspose.Cells save operation duration | C# Stopwatch profiling Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// Demonstrates using Stopwatch to record and display elapsed milliseconds for loading, recalculating formulas, autofitting columns, saving, and the total processing time of an Excel workbook with Aspose.Cells in .NET.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Overall timer for the whole process
        Stopwatch totalTimer = Stopwatch.StartNew();

        // -------------------- Load Workbook --------------------
        Stopwatch loadTimer = Stopwatch.StartNew();
        Workbook workbook = new Workbook(inputPath); // Load the workbook
        loadTimer.Stop();
        Console.WriteLine($"Load time: {loadTimer.ElapsedMilliseconds} ms");

        // -------------------- Process: Calculate Formulas --------------------
        Stopwatch calcTimer = Stopwatch.StartNew();
        workbook.CalculateFormula(); // Recalculate all formulas
        calcTimer.Stop();
        Console.WriteLine($"Formula calculation time: {calcTimer.ElapsedMilliseconds} ms");

        // -------------------- Process: AutoFit Columns (example) --------------------
        Stopwatch autofitTimer = Stopwatch.StartNew();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.AutoFitColumns(); // Adjust column widths
        autofitTimer.Stop();
        Console.WriteLine($"AutoFit columns time: {autofitTimer.ElapsedMilliseconds} ms");

        // -------------------- Save Workbook --------------------
        Stopwatch saveTimer = Stopwatch.StartNew();
        workbook.Save(outputPath, SaveFormat.Xlsx); // Save the workbook
        saveTimer.Stop();
        Console.WriteLine($"Save time: {saveTimer.ElapsedMilliseconds} ms");

        // -------------------- Total Duration --------------------
        totalTimer.Stop();
        Console.WriteLine($"Total processing time: {totalTimer.ElapsedMilliseconds} ms");
    }
}
