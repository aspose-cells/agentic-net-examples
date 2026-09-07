// Title: How to disable shared formulas for an entire workbook using Aspose.Cells for .NET by setting MaxRowsOfSharedFormula to zero
// AI Prompts: Write C# code with Aspose.Cells that sets Workbook.Settings.MaxRowsOfSharedFormula to 0 to turn off shared formulas for all worksheets. | Show a complete example that loads an existing Excel file, disables shared formulas globally via the MaxRowsOfSharedFormula property, and saves the updated workbook.
// Common Searches: Aspose.Cells C# set MaxRowsOfSharedFormula to zero to prevent shared formulas | How to turn off shared formula generation for every sheet in Aspose.Cells .NET | Disable shared formulas when saving workbook with Aspose.Cells API | Global workbook setting to stop shared formulas in Aspose.Cells for .NET
// Tags: Aspose.Cells disable shared formulas | Workbook.Settings.MaxRowsOfSharedFormula .NET | global formula sharing setting Aspose.Cells | C# Aspose.Cells workbook configuration | Excel shared formula suppression Aspose.Cells

using Aspose.Cells;

// // Creates or loads a workbook, sets Settings.MaxRowsOfSharedFormula = 0 to disable shared formulas for the whole workbook, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create rule

        // Disable shared formulas for the entire workbook
        workbook.Settings.MaxRowsOfSharedFormula = 0;

        // Save the workbook
        workbook.Save("Result.xlsx"); // save rule
    }
}
