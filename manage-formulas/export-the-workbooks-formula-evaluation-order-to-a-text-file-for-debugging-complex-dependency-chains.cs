// Title: Export Formula Evaluation Order to a Text File with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample formulas, enables the calculation chain, runs a full calculation, extracts each formula cell's precedents using GetPrecedentsInCalculation, and writes the dependency list to a plain‑text file while optionally saving the workbook.
// Keywords: Aspose.Cells C# export formula precedents | GetPrecedentsInCalculation example | calculation chain .NET | formula dependency debugging | write formula evaluation order to file | Aspose.Cells workbook analysis | C# spreadsheet formula trace
// Common Searches: Aspose.Cells retrieve formula precedents C# | export calculation chain to text file Aspose.Cells | how to get formula evaluation order .NET | debug spreadsheet dependencies with Aspose.Cells | write formula dependency list to file C#
// Developer Intent: Generate a text report of each formula cell’s precedents to help debug complex dependency chains in an Aspose.Cells workbook.
// Use Cases: Produce a readable log of formula dependencies for troubleshooting circular references or unexpected results. | Track changes in the calculation chain over time by comparing exported reports before and after formula edits. | Supply auditors or compliance teams with a plain‑text list of spreadsheet formula relationships.
// AI Prompts: Create C# code using Aspose.Cells that lists all formula cells with their precedents and saves the output as a CSV file. | Explain step‑by‑step how to enable the calculation chain in Aspose.Cells and retrieve the evaluation order for a specific worksheet. | Show how to handle cells without precedents when exporting formula dependency information to a text file in C#.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Creates a workbook, adds sample formulas, enables the calculation chain, runs a full calculation, extracts each formula cell's precedents using GetPrecedentsInCalculation, and writes the dependency list to a plain‑text file while optionally saving the workbook.
class ExportFormulaEvaluationOrder
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample data and formulas to demonstrate dependencies
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["B1"].Formula = "=A1+A2";          // B1 depends on A1 and A2
        cells["B2"].Formula = "=B1*2";           // B2 depends on B1
        cells["C1"].Formula = "=SUM(A1:A2)";     // C1 depends on A1 and A2
        cells["C2"].Formula = "=B2+C1";          // C2 depends on B2 and C1

        // Enable the calculation chain so dependency information can be retrieved
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Perform a full calculation to build the chain
        workbook.CalculateFormula();

        // Export the evaluation order (precedents) to a text file
        using (StreamWriter writer = new StreamWriter("FormulaEvaluationOrder.txt"))
        {
            // Iterate through all cells in the worksheet
            foreach (Cell cell in cells)
            {
                // Process only cells that contain formulas
                if (cell.IsFormula)
                {
                    writer.WriteLine($"Cell {cell.Name} depends on:");

                    // Get the precedents for the current formula cell
                    IEnumerator precedents = cell.GetPrecedentsInCalculation();

                    if (precedents != null)
                    {
                        while (precedents.MoveNext())
                        {
                            // Each item is a ReferredArea describing a range of precedent cells
                            ReferredArea area = (ReferredArea)precedents.Current;
                            writer.WriteLine($"  {area}");
                        }
                    }
                    else
                    {
                        writer.WriteLine("  (no precedents)");
                    }
                }
            }
        }

        // Optionally save the workbook for reference
        workbook.Save("SampleWorkbook.xlsx");
    }
}
