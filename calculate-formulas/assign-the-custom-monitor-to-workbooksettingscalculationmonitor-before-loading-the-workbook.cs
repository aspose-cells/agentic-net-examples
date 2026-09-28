// Title: Assign a custom AbstractCalculationMonitor to CalculationOptions before opening a workbook with Aspose.Cells for .NET
// AI Prompts: Create a subclass of AbstractCalculationMonitor that logs before/after each cell calculation and detects circular references, assign it to CalculationOptions.CalculationMonitor, then load the workbook and invoke CalculateFormula. | Implement OnCircular in a custom calculation monitor to abort formula evaluation on circular references, set the monitor via CalculationOptions before constructing the Workbook, and run the calculation.
// Common Searches: how to use a custom calculation monitor with Aspose.Cells calculation options in C# | Aspose.Cells calculate formulas while logging each cell evaluation | prevent circular reference errors in Aspose.Cells by using a calculation monitor | set CalculationOptions.CalculationMonitor before Workbook constructor Aspose.Cells | log before and after cell calculation events using Aspose.Cells .NET API
// Tags: custom AbstractCalculationMonitor Aspose.Cells | CalculationOptions.CalculationMonitor usage | log cell calculation events Aspose.Cells | circular reference handling Aspose.Cells | calculate formulas with custom monitor .NET

using System;
using System.Collections;
using Aspose.Cells;

// Defines a CustomCalculationMonitor derived from AbstractCalculationMonitor to log cell calculations and detect circular references, assigns it to CalculationOptions before loading the workbook, runs CalculateFormula with the monitor, and saves the updated file.
public class CustomCalculationMonitor : AbstractCalculationMonitor
{
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        Console.WriteLine($"Before calculating cell - Sheet: {sheetIndex}, Row: {rowIndex}, Column: {columnIndex}");
    }

    public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        Console.WriteLine($"After calculating cell - Sheet: {sheetIndex}, Row: {rowIndex}, Column: {columnIndex}");
        Console.WriteLine($"Value changed: {ValueChanged}, Original: {OriginalValue}, New: {CalculatedValue}");
    }

    public override bool OnCircular(IEnumerator circularCellsData)
    {
        Console.WriteLine("Circular reference detected.");
        // Return false to stop calculation when a circular reference is found
        return false;
    }
}

public class Program
{
    public static void Main()
    {
        // Create calculation options and assign the custom monitor BEFORE loading the workbook
        CalculationOptions calcOptions = new CalculationOptions
        {
            CalculationMonitor = new CustomCalculationMonitor()
        };

        // Load the workbook (the monitor will be used later during calculation)
        Workbook workbook = new Workbook("input.xlsx");

        // Perform formula calculation using the options that contain the monitor
        workbook.CalculateFormula(calcOptions);

        // Save the workbook after calculation
        workbook.Save("output.xlsx");
    }
}
