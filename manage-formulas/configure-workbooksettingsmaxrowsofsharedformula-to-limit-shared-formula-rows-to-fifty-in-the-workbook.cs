// Title: How to limit shared formula rows to 50 in an Aspose.Cells workbook using C#
// AI Prompts: Set workbook.Settings.MaxRowsOfSharedFormula = 50 and save the workbook with Aspose.Cells in C#. | Configure the Aspose.Cells workbook to restrict shared formula rows to fifty before exporting to Excel.
// Common Searches: Aspose.Cells C# limit number of rows for shared formulas to 50 | Set MaxRowsOfSharedFormula property in Aspose.Cells workbook | Restrict shared formula rows count when generating Excel with Aspose.Cells | How to configure workbook settings to cap shared formula rows in C#
// Tags: Aspose.Cells workbook Settings MaxRowsOfSharedFormula | C# limit shared formula rows Excel | configure shared formula row limit Aspose.Cells | set shared formula max rows Aspose.Cells C#

using Aspose.Cells;

// The example creates a new Workbook, sets Settings.MaxRowsOfSharedFormula to 50 to cap shared formula rows, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Limit shared formula rows to fifty
        workbook.Settings.MaxRowsOfSharedFormula = 50;

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
