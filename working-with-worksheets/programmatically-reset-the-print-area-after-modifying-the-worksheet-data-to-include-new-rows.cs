// Title: Reset the Excel worksheet print area after adding rows with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear worksheet print area after inserting rows | programmatically reset PageSetup.PrintArea in .NET after adding data | how to recalculate Excel print area when rows are added using Aspose | C# Aspose.Cells remove existing print area before saving workbook | update Excel print range after dynamic row insertion with Aspose.Cells
// Tags: Aspose.Cells reset print area C# | clear PageSetup.PrintArea after row insertion | modify worksheet print settings .NET | insert rows and update print range Aspose | Excel print area recalculation using Aspose.Cells

using Aspose.Cells;

// Loads an Excel workbook, inserts additional rows after the last used row, clears the worksheet's print area by setting PageSetup.PrintArea to an empty string, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert new rows after the last used row (example: add 5 rows)
        int lastRow = sheet.Cells.MaxDataRow;
        sheet.Cells.InsertRows(lastRow + 1, 5);

        // Reset the print area so it will be recalculated based on the new data
        sheet.PageSetup.PrintArea = "";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
