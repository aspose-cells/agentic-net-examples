// Title: How to limit shared formula rows to 50 using Aspose.Cells Workbook.Settings.MaxRowsOfSharedFormula in C#
// Common Searches: Aspose.Cells C# set maximum rows for shared formulas | How to restrict shared formula rows to 50 in a .NET workbook | Workbook.Settings.MaxRowsOfSharedFormula example code | Limit number of rows that can use shared formulas with Aspose.Cells | C# Excel file shared formula row limit using Aspose.Cells
// Tags: Workbook.Settings.MaxRowsOfSharedFormula usage | shared formula capacity setting Aspose.Cells | C# Excel workbook formula row restriction | Aspose.Cells configure formula settings | limit shared formula rows programmatically

using Aspose.Cells;

// The example creates a new Workbook, sets the MaxRowsOfSharedFormula property to 50 to cap the number of rows that can contain shared formulas, and saves the file as SharedFormulaLimited.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Limit shared formula rows to fifty
        workbook.Settings.MaxRowsOfSharedFormula = 50;

        // Save the workbook
        workbook.Save("SharedFormulaLimited.xlsx");
    }
}
