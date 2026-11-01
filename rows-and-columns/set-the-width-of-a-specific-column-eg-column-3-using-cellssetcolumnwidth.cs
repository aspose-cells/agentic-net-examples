// Title: Set the width of column 3 (zero‑based index) to 20 characters using Aspose.Cells SetColumnWidth in C#
// Common Searches: how to set column width for column 3 using Aspose.Cells in C# | Aspose.Cells SetColumnWidth example with zero based column index | C# code to change Excel column width to a specific character count using Aspose.Cells | saving a workbook after modifying column width with Aspose.Cells | set column width in a newly created workbook using Aspose.Cells C#
// Tags: Aspose.Cells column width adjustment | SetColumnWidth zero based index | Excel column width in characters Aspose.Cells | modify column width before saving workbook | C# workbook creation column width Aspose.Cells

using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, retrieves its Cells collection, sets the width of the third column (zero‑based index 2) to 20 characters with SetColumnWidth, and saves the file as ColumnWidthExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the cells collection
        Cells cells = worksheet.Cells;

        // Set the width of column 3 (zero‑based index 2) to 20 characters
        cells.SetColumnWidth(2, 20);

        // Save the workbook (using the save rule)
        workbook.Save("ColumnWidthExample.xlsx");
    }
}
