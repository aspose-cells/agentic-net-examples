// Title: Set the maximum rows for shared formula blocks to 100 in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# set MaxRowsOfSharedFormula to 100 | increase shared formula row span in Aspose.Cells workbook | how to configure shared formula size limit with Aspose.Cells API | C# workbook settings for larger shared formula blocks Aspose.Cells | adjust MaxRowsOfSharedFormula property in Aspose.Cells
// Tags: MaxRowsOfSharedFormula property configuration | shared formula row limit Aspose.Cells | workbook settings for shared formulas C# | increase shared formula block size Aspose.Cells | Aspose.Cells workbook max rows shared formula

using Aspose.Cells;

// Creates a new Aspose.Cells Workbook, sets Settings.MaxRowsOfSharedFormula to 100 to allow larger shared formula blocks, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Set the maximum number of rows that a shared formula can span
        workbook.Settings.MaxRowsOfSharedFormula = 100;

        // Save the workbook to verify the setting (optional)
        workbook.Save("Result.xlsx");
    }
}
