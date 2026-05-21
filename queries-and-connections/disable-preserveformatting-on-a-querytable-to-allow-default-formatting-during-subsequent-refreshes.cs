using System;
using Aspose.Cells;

class DisableQueryTablePreserveFormatting
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data (optional, just to have a sheet with content)
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("John");

        // QueryTables are usually created from an external data source.
        // Here we simply check if any QueryTable already exists.
        if (worksheet.QueryTables.Count > 0)
        {
            // Get the first QueryTable
            QueryTable queryTable = worksheet.QueryTables[0];

            // Disable PreserveFormatting so that default formatting is applied on refresh
            queryTable.PreserveFormatting = false;

            Console.WriteLine("QueryTable PreserveFormatting set to false.");
        }
        else
        {
            Console.WriteLine("No QueryTable found in the worksheet.");
        }

        // Save the workbook (lifecycle: save)
        workbook.Save("DisablePreserveFormatting.xlsx");
    }
}