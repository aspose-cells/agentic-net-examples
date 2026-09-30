// Title: How to turn off PreserveFormatting for a QueryTable in Aspose.Cells for .NET to use default formatting on refresh
// AI Prompts: Set queryTable.PreserveFormatting = false and save the workbook with Aspose.Cells. | Iterate over every QueryTable in a worksheet, disable PreserveFormatting, then refresh the table using C#. | Modify an existing Excel file so that its QueryTable applies default cell styles after each refresh with Aspose.Cells.
// Common Searches: aspocells c# turn off preserveformatting for querytable before refresh | how to make querytable use default styles after refresh using Aspose.Cells | c# Aspose.Cells example disabling formatting preservation on Excel querytable | remove formatting lock on querytable refresh Aspose.Cells .NET
// Tags: aspocells querytable preserveformatting flag | disable formatting preservation on querytable .net | default cell style on querytable refresh aspocells | c# modify querytable properties workbook | excel querytable refresh formatting control aspocells

using Aspose.Cells;
using System;

// The sample loads an Excel workbook, accesses the first worksheet, checks for a QueryTable, disables its PreserveFormatting property so that default formatting is applied on subsequent refreshes, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet containing the QueryTable
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one QueryTable
        if (sheet.QueryTables.Count > 0)
        {
            // Get the first QueryTable (adjust index as needed)
            QueryTable queryTable = sheet.QueryTables[0];

            // Disable PreserveFormatting so default formatting is applied on refresh
            queryTable.PreserveFormatting = false;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
