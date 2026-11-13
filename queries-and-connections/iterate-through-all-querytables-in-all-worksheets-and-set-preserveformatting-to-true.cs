// Title: Set PreserveFormatting = true for all QueryTables across every worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set PreserveFormatting on all QueryTables in a workbook | how to enable formatting preservation for each QueryTable when processing Excel files with Aspose.Cells | iterate through worksheets and QueryTables to apply PreserveFormatting using Aspose.Cells .NET API | C# code to preserve Excel query table formatting after import with Aspose.Cells | bulk update QueryTable PreserveFormatting property in multiple sheets using Aspose.Cells
// Tags: querytable preserveformatting aspnet | aspose.cells bulk querytable formatting | excel workbook iterate querytables c# | aspose.cells set querytable formatting property | c# preserve querytable style aspnet

using Aspose.Cells;

// The example loads an existing Excel workbook, loops through each worksheet and each QueryTable within those sheets, sets the QueryTable's PreserveFormatting property to true, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through each QueryTable in the current worksheet
            foreach (QueryTable queryTable in sheet.QueryTables)
            {
                // Enable PreserveFormatting for the QueryTable
                queryTable.PreserveFormatting = true;
            }
        }

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
