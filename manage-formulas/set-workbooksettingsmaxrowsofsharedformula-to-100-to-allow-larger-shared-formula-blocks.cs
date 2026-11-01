// Title: Set MaxRowsOfSharedFormula to 100 in Aspose.Cells for .NET to enable larger shared formula blocks
// Common Searches: how to change MaxRowsOfSharedFormula setting in Aspose.Cells C# | increase the row limit for shared formulas in an Excel workbook using Aspose.Cells | Aspose.Cells workbook settings for larger shared formula blocks | C# code to set shared formula block size to 100 rows with Aspose.Cells | adjust shared formula row count limit in Aspose.Cells .NET API
// Tags: MaxRowsOfSharedFormula property Aspose.Cells | shared formula row limit configuration .NET | workbook settings for shared formulas C# | increase shared formula block size Aspose.Cells | adjust shared formula rows in Excel workbook

using Aspose.Cells;

// The example creates a new Workbook, sets Settings.MaxRowsOfSharedFormula to 100 to allow larger shared formula blocks, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the maximum number of rows for a shared formula block to 100
        workbook.Settings.MaxRowsOfSharedFormula = 100;

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
