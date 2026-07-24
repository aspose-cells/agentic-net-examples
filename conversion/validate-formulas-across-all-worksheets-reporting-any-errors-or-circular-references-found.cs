// Title: Validate Excel Formulas and Detect Circular References with Aspose.Cells for .NET (C#)
// Description: Load an Excel workbook, enable the calculation chain, and configure CalculationOptions to surface errors. The sample runs workbook.CalculateFormula with a custom CircularReferenceMonitor, catches exceptions, scans every used cell across all worksheets for CellValueType.IsError, and prints each error address and message. Finally, the workbook can be saved after corrections.
// Keywords: Aspose.Cells formula validation | C# Excel circular reference detection | Aspose.Cells CalculateFormula error handling | scan workbook for formula errors .NET | AbstractCalculationMonitor example | Excel error values C# | calculation chain Aspose.Cells | US developers Aspose.Cells | UK .NET Excel automation | India C# spreadsheet processing
// Common Searches: How to find circular references in Excel using Aspose.Cells C# | List cells with formula errors after calculation Aspose.Cells | Enable calculation chain and capture errors Aspose.Cells .NET | Iterate worksheets to detect #REF! or #DIV/0! in C# | Aspose.Cells example for formula validation and error reporting
// Developer Intent: Run a comprehensive formula audit on an Excel file and output any error cells or circular‑reference chains before saving.
// Use Cases: Execute workbook.CalculateFormula with a CircularReferenceMonitor to expose circular dependencies. | After calculation, loop through each worksheet’s used range and log cells where Cell.Type equals CellValueType.IsError. | Generate a report of error locations for QA or automated correction workflows. | Save the corrected workbook to a new file after fixing reported issues.
// AI Prompts: Write C# code using Aspose.Cells that validates all formulas in a workbook and returns a list of cells containing errors. | Create a custom AbstractCalculationMonitor that logs each step of a circular reference chain and allows the engine to continue processing. | Show how to set CalculationOptions so that errors are not ignored and circular references are captured during formula calculation.

using System;
using System.Collections;
using Aspose.Cells;

// Load an Excel workbook, enable the calculation chain, and configure CalculationOptions to surface errors. The sample runs workbook.CalculateFormula with a custom CircularReferenceMonitor, catches exceptions, scans every used cell across all worksheets for CellValueType.IsError, and prints each error address and message. Finally, the workbook can be saved after corrections.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable calculation chain to allow dependent analysis
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Configure calculation options
        CalculationOptions options = new CalculationOptions
        {
            // Do not ignore errors so they will be reported
            IgnoreError = false,
            // Attach a monitor to capture circular references
            CalculationMonitor = new CircularReferenceMonitor()
        };

        // Perform formula calculation and capture any errors
        try
        {
            workbook.CalculateFormula(options);
            Console.WriteLine("All formulas calculated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during formula calculation: " + ex.Message);
        }

        // After calculation, scan all cells for error values
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            // Iterate only over the used range to improve performance
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];
                    // Cells that resulted in an error have the IsError type
                    if (cell.Type == CellValueType.IsError)
                    {
                        Console.WriteLine($"Error in {sheet.Name}!{cell.Name}: {cell.StringValue}");
                    }
                }
            }
        }

        // Save the workbook (optional, e.g., after fixing issues)
        workbook.Save("output.xlsx");
    }

    // Custom monitor to detect circular references during calculation
    private class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell representing a cell involved in the circular reference
                Console.WriteLine(circularCellsData.Current);
            }
            // Return true to let the engine continue processing other cells
            return true;
        }
    }
}
