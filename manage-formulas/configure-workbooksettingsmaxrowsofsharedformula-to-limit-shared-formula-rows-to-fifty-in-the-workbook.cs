// Title: Set Workbook.Settings.MaxRowsOfSharedFormula to 50 in Aspose.Cells (C#)
// Description: Creates a new Workbook, configures Workbook.Settings.MaxRowsOfSharedFormula to 50, applies a shared formula intended for 100 rows (only the first 50 are filled), and saves the result as MaxRowsOfSharedFormulaDemo.xlsx.
// Keywords: Aspose.Cells | MaxRowsOfSharedFormula | C# | shared formula limit | Workbook.Settings | SetSharedFormula example | Excel performance optimization | limit formula rows | Aspose.Cells tutorial
// Common Searches: Aspose.Cells set MaxRowsOfSharedFormula property | limit shared formula rows C# | how to restrict shared formula expansion in Aspose.Cells | MaxRowsOfSharedFormula example for 50 rows | shared formula row cap Aspose.Cells
// Developer Intent: Restrict the number of rows a shared formula can populate to 50 within a workbook.
// Use Cases: Prevent excessive memory usage when generating large reports. | Maintain compatibility with older Excel versions that have row limits for shared formulas. | Control file size by capping formula propagation in automated spreadsheet creation.
// AI Prompts: Show code to verify that exactly 50 rows contain the shared formula after applying it to 100 rows. | Generate a version that reads the row limit from an appsettings.json file. | Explain how MaxRowsOfSharedFormula interacts with the rowCount argument of SetSharedFormula.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Creates a new Workbook, configures Workbook.Settings.MaxRowsOfSharedFormula to 50, applies a shared formula intended for 100 rows (only the first 50 are filled), and saves the result as MaxRowsOfSharedFormulaDemo.xlsx.
    class MaxRowsOfSharedFormulaExample
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Configure the workbook to limit shared formula rows to 50
            workbook.Settings.MaxRowsOfSharedFormula = 50;

            // (Optional) Demonstrate the effect by setting a shared formula that exceeds the limit
            // This will only populate up to 50 rows despite the request for more
            Worksheet sheet = workbook.Worksheets[0];
            Cell startCell = sheet.Cells["B1"];
            // Attempt to set a shared formula for 100 rows; only 50 will be applied due to the limit
            startCell.SetSharedFormula("=A1", 100, 1);

            // Save the workbook (lifecycle: save)
            workbook.Save("MaxRowsOfSharedFormulaDemo.xlsx");

            Console.WriteLine("Workbook saved with MaxRowsOfSharedFormula set to 50.");
        }
    }
}
