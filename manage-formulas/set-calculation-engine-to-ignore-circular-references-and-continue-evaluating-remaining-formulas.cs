// Title: Ignore Circular References and Continue Formula Evaluation with Aspose.Cells for .NET
// Description: Demonstrates how to configure Aspose.Cells CalculationOptions with a custom AbstractCalculationMonitor that logs circular cells and returns true, allowing the engine to skip the circular loop and still calculate independent formulas such as C1 = A1+10. The workbook is then saved with the results.
// Keywords: Aspose.Cells | C# | .NET | circular reference | ignore circular reference | continue calculation | CalculationOptions | AbstractCalculationMonitor | custom monitor | workbook.CalculateFormula
// Common Searches: Aspose.Cells skip circular reference | C# ignore circular reference during calculation | continue formula evaluation after circular reference Aspose.Cells | how to use AbstractCalculationMonitor in Aspose.Cells | set calculation options to ignore circular loops .NET
// Developer Intent: Skip circular references while still evaluating other formulas in a workbook.
// Use Cases: Log circular cells without aborting the calculation process. | Process large spreadsheets that contain intentional circular formulas while retrieving values from independent cells. | Integrate a custom monitor to keep calculations alive in automated reporting pipelines.
// AI Prompts: Generate C# code that creates a CalculationOptions with an AbstractCalculationMonitor which records circular cell addresses and returns true to continue calculation. | Show how to modify the CircularReferenceMonitor to store detected circular references in a collection for later analysis. | Explain how to configure Aspose.Cells to treat circular references as zero while still calculating non‑circular cells.

using System;
using System.Collections;
using Aspose.Cells;

// Demonstrates how to configure Aspose.Cells CalculationOptions with a custom AbstractCalculationMonitor that logs circular cells and returns true, allowing the engine to skip the circular loop and still calculate independent formulas such as C1 = A1+10. The workbook is then saved with the results.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set up a circular reference: A1 -> B1 -> A1
        sheet.Cells["A1"].Formula = "=B1";
        sheet.Cells["B1"].Formula = "=A1";

        // Add another formula that depends on the circular cells
        // This will be evaluated after the circular reference is handled
        sheet.Cells["C1"].Formula = "=A1+10";

        // Create calculation options and attach a custom monitor
        CalculationOptions calcOptions = new CalculationOptions
        {
            CalculationMonitor = new CircularReferenceMonitor()
        };

        // Perform calculation; the monitor will handle circular references
        workbook.CalculateFormula(calcOptions);

        // Display the results
        Console.WriteLine($"A1 value: {sheet.Cells["A1"].Value}");
        Console.WriteLine($"B1 value: {sheet.Cells["B1"].Value}");
        Console.WriteLine($"C1 value: {sheet.Cells["C1"].Value}");

        // Save the workbook
        workbook.Save("CircularHandled.xlsx");
    }

    // Custom monitor that logs circular references and tells the engine to continue
    private class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                Console.WriteLine($"  {circularCellsData.Current}");
            }
            // Return true to let the calculation engine continue processing other cells
            return true;
        }
    }
}
