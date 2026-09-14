// Title: Disable PreserveFormatting for a QueryTable in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to set QueryTable.PreserveFormatting to false and save the workbook. | Load a workbook, find the first QueryTable, turn off its PreserveFormatting property, and write the changes back. | Write C# code that iterates over all QueryTables in a worksheet and disables PreserveFormatting with Aspose.Cells.
// Common Searches: Aspose.Cells C# disable PreserveFormatting on QueryTable | How to turn off formatting preservation when refreshing an Excel QueryTable with Aspose.Cells | Set QueryTable PreserveFormatting false programmatically .NET | Refresh Excel query table without keeping original formatting using Aspose.Cells
// Tags: Aspose.Cells QueryTable PreserveFormatting property | disable PreserveFormatting for Excel QueryTable .NET | set QueryTable.PreserveFormatting false C# | refresh QueryTable without preserving formatting Aspose.Cells | modify QueryTable settings programmatically C#

using Aspose.Cells;
using System;

// Loads an existing workbook, accesses the first worksheet, disables the PreserveFormatting flag of the first QueryTable, outputs the new setting, and saves the workbook as a new file.
class DisableQueryTablePreserveFormatting
{
    static void Main()
    {
        // Load an existing workbook that contains a query table
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Check if the worksheet has any query tables
        if (worksheet.QueryTables.Count > 0)
        {
            // Get the first query table
            QueryTable queryTable = worksheet.QueryTables[0];

            // Disable PreserveFormatting so default formatting is applied on refresh
            queryTable.PreserveFormatting = false;

            Console.WriteLine("PreserveFormatting set to: " + queryTable.PreserveFormatting);
        }
        else
        {
            Console.WriteLine("No query tables found in the worksheet.");
        }

        // Save the workbook with the updated setting
        workbook.Save("output.xlsx");
    }
}
