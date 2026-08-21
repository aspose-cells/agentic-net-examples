// Title: Aspose.Cells C# – Detect Circular References and Export Dependent Formulas to a Report Sheet
// Description: This example shows how to load or create a workbook, attach a custom AbstractCalculationMonitor that captures the addresses of cells involved in circular references, run workbook.CalculateFormula, retrieve each circular cell's dependents with GetDependentsInCalculation(true), and write the circular cell, dependent cell, and dependent formula to a new worksheet named "CircularReferenceReport" before saving the file.
// Keywords: Aspose.Cells circular reference detection | C# calculation monitor | GetDependentsInCalculation | Excel dependent formulas report | .NET Excel circular reference | Aspose.Cells API example
// Common Searches: how to capture circular references with Aspose.Cells | list cells that depend on a circular reference in C# | generate circular reference report worksheet Aspose.Cells | custom calculation monitor example .NET | retrieve dependent formulas from circular cells
// Developer Intent: Find circular‑reference cells, enumerate their dependent formulas, and produce an Excel summary sheet.
// Use Cases: Attach a custom AbstractCalculationMonitor to record circular cell addresses during workbook.CalculateFormula. | Use GetDependentsInCalculation(true) to collect all cells whose results rely on each circular cell. | Write the circular cell address, dependent cell address, and dependent formula to a dedicated report worksheet for analysis. | Save the workbook with the report for auditing or debugging complex Excel models.
// AI Prompts: Generate C# code using Aspose.Cells that detects circular references, lists dependent cells, and creates a report worksheet. | Show how to implement a custom CalculationMonitor to capture circular cell addresses and then export dependent formulas to a new sheet. | Provide an Aspose.Cells example that retrieves dependents of a circular cell with GetDependentsInCalculation and saves the results as an Excel report.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// This example shows how to load or create a workbook, attach a custom AbstractCalculationMonitor that captures the addresses of cells involved in circular references, run workbook.CalculateFormula, retrieve each circular cell's dependents with GetDependentsInCalculation(true), and write the circular cell, dependent cell, and dependent formula to a new worksheet named "CircularReferenceReport" before saving the file.
class CircularReferenceMonitor : AbstractCalculationMonitor
{
    // Stores addresses of cells that are part of circular references
    public List<string> CircularCellAddresses { get; } = new List<string>();

    // Called by the calculation engine when a circular reference is detected
    public override bool OnCircular(IEnumerator circularCellsData)
    {
        while (circularCellsData.MoveNext())
        {
            // Each item is a Cell (or CalculationCell) involved in the circular reference.
            // Retrieve its address via the Name property.
            string address = string.Empty;
            if (circularCellsData.Current is Cell cell)
            {
                address = cell.Name;
            }
            else if (circularCellsData.Current != null)
            {
                // Fallback to string representation if casting fails.
                address = circularCellsData.Current.ToString();
            }

            if (!string.IsNullOrEmpty(address))
            {
                CircularCellAddresses.Add(address);
            }
        }
        // Continue calculation for these cells.
        return true;
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a workbook (or load an existing one)
            // -------------------------------------------------
            Workbook workbook = new Workbook(); // create new workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data that creates a circular reference:
            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";

            // -------------------------------------------------
            // 2. Prepare calculation options with a custom monitor
            // -------------------------------------------------
            CalculationOptions options = new CalculationOptions();
            CircularReferenceMonitor monitor = new CircularReferenceMonitor();
            options.CalculationMonitor = monitor;

            // Ensure iterative calculation is disabled (default behavior)
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = false;

            // -------------------------------------------------
            // 3. Perform calculation (circular references will trigger the monitor)
            // -------------------------------------------------
            workbook.CalculateFormula(options);

            // -------------------------------------------------
            // 4. Build a report worksheet with circular cells and their dependents
            // -------------------------------------------------
            Worksheet report = workbook.Worksheets.Add("CircularReferenceReport");
            int reportRow = 0;
            report.Cells[reportRow, 0].PutValue("Circular Cell");
            report.Cells[reportRow, 1].PutValue("Dependent Cell");
            report.Cells[reportRow, 2].PutValue("Dependent Formula");
            reportRow++;

            foreach (string circAddress in monitor.CircularCellAddresses)
            {
                Cell circCell = sheet.Cells[circAddress];

                // Get all cells whose calculated result depends on the circular cell
                IEnumerator dependents = circCell.GetDependentsInCalculation(true);

                if (dependents != null)
                {
                    while (dependents.MoveNext())
                    {
                        if (dependents.Current is Cell depCell)
                        {
                            report.Cells[reportRow, 0].PutValue(circAddress);
                            report.Cells[reportRow, 1].PutValue(depCell.Name);
                            report.Cells[reportRow, 2].PutValue(depCell.Formula);
                            reportRow++;
                        }
                    }
                }
                else
                {
                    // No dependents found; just list the circular cell
                    report.Cells[reportRow, 0].PutValue(circAddress);
                    reportRow++;
                }
            }

            // -------------------------------------------------
            // 5. Save the workbook (including the report)
            // -------------------------------------------------
            string outputPath = "CircularReferenceReport.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Report saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
