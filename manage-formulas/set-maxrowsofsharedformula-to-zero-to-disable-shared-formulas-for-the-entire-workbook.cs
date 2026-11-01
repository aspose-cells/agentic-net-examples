// Title: How to disable shared formulas for an entire workbook by setting MaxRowsOfSharedFormula to 0 with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set MaxRowsOfSharedFormula to zero to turn off shared formulas | C# disable shared formulas for all worksheets in Aspose.Cells | How to prevent shared formula generation when saving a workbook using Aspose.Cells | Aspose.Cells workbook Settings.MaxRowsOfSharedFormula example in .NET | Turn off shared formulas globally in Aspose.Cells .NET
// Tags: Settings.MaxRowsOfSharedFormula property | disable shared formulas workbook | global formula sharing Aspose.Cells | Aspose.Cells workbook settings C# | shared formula limit zero Excel

using Aspose.Cells;
using System;

// Creates (or loads) a workbook, sets Settings.MaxRowsOfSharedFormula to 0 to disable shared formulas for the whole workbook, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Disable shared formulas for the entire workbook
        // Setting MaxRowsOfSharedFormula to 0 disables shared formulas
        workbook.Settings.MaxRowsOfSharedFormula = 0;

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
