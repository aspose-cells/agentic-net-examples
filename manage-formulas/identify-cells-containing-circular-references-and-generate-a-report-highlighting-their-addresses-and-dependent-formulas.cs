// Title: Aspose.Cells .NET: Detect Circular References and Generate a Report
// Description: Loads an Excel workbook, sets a CalculationOptions with a custom CircularReferenceMonitor, captures every cell involved in a circular reference and its dependent formulas, writes the findings to a new worksheet named "CircularReport", and saves the updated file.
// Keywords: Aspose.Cells circular reference detection | C# Excel circular reference monitor | Aspose.Cells CalculationOptions | Excel circular formula report | GetDependents Aspose.Cells | detect circular formulas .NET | Excel diagnostics Aspose | US .NET developers | European C# Excel automation
// Common Searches: how to log circular reference cells with Aspose.Cells | generate circular reference report in C# | Aspose.Cells AbstractCalculationMonitor example | list dependent formulas for circular cells | Excel circular reference detection .NET library
// Developer Intent: Find cells that cause circular calculations, collect their addresses and dependent formulas, and output the information to a dedicated worksheet.
// Use Cases: Audit complex financial models for circular logic and provide a clear remediation list. | Validate user‑uploaded spreadsheets on a server, rejecting files that contain circular formulas before processing. | Supply end‑users with an automatic diagnostic sheet that pinpoints problematic cells and shows related formulas.
// AI Prompts: Create a function that returns circular reference details as a JSON array instead of writing to a worksheet. | Enhance the monitor to include the recursion depth of each dependent formula in the report. | Write a unit test that confirms CircularReferenceMonitor captures all circular cells and their dependents correctly.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook, sets a CalculationOptions with a custom CircularReferenceMonitor, captures every cell involved in a circular reference and its dependent formulas, writes the findings to a new worksheet named "CircularReport", and saves the updated file.
class Program
{
    // Holds report entries: "CellAddress|DependentFormulas"
    static List<string> reportLines = new List<string>();

    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set up calculation options with a custom monitor to capture circular references
        CalculationOptions options = new CalculationOptions();
        options.CalculationMonitor = new CircularReferenceMonitor();

        // Perform formula calculation; the monitor will be invoked for circular references
        workbook.CalculateFormula(options);

        // Create a worksheet to hold the circular reference report
        Worksheet reportSheet = workbook.Worksheets.Add("CircularReport");
        int row = 0;
        reportSheet.Cells[row, 0].PutValue("Circular Cell");
        reportSheet.Cells[row, 1].PutValue("Dependent Formulas");
        row++;

        // Populate the report sheet with collected data
        foreach (string line in reportLines)
        {
            string[] parts = line.Split('|');
            reportSheet.Cells[row, 0].PutValue(parts[0]);               // Cell address
            reportSheet.Cells[row, 1].PutValue(parts.Length > 1 ? parts[1] : ""); // Dependents
            row++;
        }

        // Save the workbook with the added report
        workbook.Save("output_with_circular_report.xlsx");
    }

    // Custom calculation monitor that captures circular reference information
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            // Iterate over each cell involved in the circular reference
            while (circularCellsData.MoveNext())
            {
                object current = circularCellsData.Current;

                // Try to treat the item as a Cell; if not possible, fall back to its string representation
                Cell cell = current as Cell;
                string address = cell != null ? cell.Name : current?.ToString() ?? "Unknown";
                string formula = cell != null ? cell.Formula : string.Empty;

                // Gather dependent cells (recursive) if we have a valid Cell object
                string dependentsInfo = string.Empty;
                if (cell != null)
                {
                    Cell[] dependents = cell.GetDependents(true);
                    List<string> depList = new List<string>();
                    foreach (Cell dep in dependents)
                    {
                        depList.Add($"{dep.Name}:{dep.Formula}");
                    }
                    dependentsInfo = string.Join(", ", depList);
                }

                // Store the information for later reporting
                Program.reportLines.Add($"{address}|{dependentsInfo}");
            }

            // Return true to let the engine continue processing other cells
            return true;
        }
    }
}
