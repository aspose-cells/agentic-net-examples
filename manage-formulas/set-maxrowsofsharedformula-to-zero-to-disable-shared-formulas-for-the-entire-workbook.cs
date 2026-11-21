// Title: Disable shared formulas for an entire workbook by setting MaxRowsOfSharedFormula to zero in Aspose.Cells for .NET
// Common Searches: Aspose.Cells how to disable shared formulas for all worksheets in C# | set workbook.Settings.MaxRowsOfSharedFormula to 0 using Aspose.Cells .NET | turn off formula sharing across a workbook with Aspose.Cells API | C# code to prevent shared formulas when saving Excel file with Aspose.Cells | disable shared formula optimization in Aspose.Cells workbook settings
// Tags: MaxRowsOfSharedFormula setting | disable shared formulas workbook | Aspose.Cells workbook settings C# | formula sharing optimization off | Excel file without shared formulas Aspose.Cells

using Aspose.Cells;

// Creates (or loads) a workbook, sets Settings.MaxRowsOfSharedFormula to 0 to turn off shared formulas for the entire workbook, and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Disable shared formulas for the entire workbook
        // Setting MaxRowsOfSharedFormula to 0 turns off shared formulas.
        workbook.Settings.MaxRowsOfSharedFormula = 0;

        // Save the workbook to a file
        workbook.Save("DisabledSharedFormulas.xlsx");
    }
}
