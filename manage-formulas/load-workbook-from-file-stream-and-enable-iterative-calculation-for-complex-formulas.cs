// Title: Load Excel Workbook from FileStream and Enable Iterative Calculation with Aspose.Cells for .NET
// Description: Shows how to open an .xlsx file via FileStream, load it into an Aspose.Cells Workbook, activate iterative calculation to resolve circular references, configure MaxIteration and MaxChange, recalculate formulas, and save the updated file.
// Keywords: Aspose.Cells FileStream load | iterative calculation .NET | circular reference handling Aspose.Cells | MaxIteration Aspose.Cells | MaxChange Aspose.Cells | calculate formulas Aspose.Cells | Workbook.Save C# | LoadOptions stream Aspose.Cells
// Common Searches: load excel from filestream aspose.cells | enable iterative calculation aspose.cells .net | set maxiteration and maxchange aspose.cells | handle circular references in aspose.cells workbook | recalculate formulas after enabling iterative calculation
// Developer Intent: Load a workbook from a stream, turn on iterative calculation with custom limits, recalculate formulas, and save the file.
// Use Cases: Process uploaded Excel files received as streams while automatically resolving circular references. | Run financial or engineering models that require iterative formula evaluation with specific iteration thresholds. | Expose a web API that reads Excel data from a stream, enables iterative calculation, and returns the modified workbook.
// AI Prompts: Generate C# code to load an Excel file from a MemoryStream, enable iterative calculation, and set MaxIteration and MaxChange using Aspose.Cells. | Explain how to configure iterative calculation parameters in Aspose.Cells and recalculate all formulas after loading a workbook from a stream. | Provide best‑practice recommendations for handling circular references with Aspose.Cells iterative calculation in .NET applications.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to open an .xlsx file via FileStream, load it into an Aspose.Cells Workbook, activate iterative calculation to resolve circular references, configure MaxIteration and MaxChange, recalculate formulas, and save the updated file.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Open the file as a stream
        using (FileStream fileStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        {
            // Create load options (default settings)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook from the stream using the constructor that accepts a Stream and LoadOptions
            Workbook workbook = new Workbook(fileStream, loadOptions);

            // Enable iterative calculation to resolve circular references or complex formulas
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;   // maximum change allowed per iteration

            // Optionally calculate all formulas after enabling iterative calculation
            workbook.CalculateFormula();

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
