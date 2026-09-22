// Title: Delete a specific PivotTable from an Excel worksheet by index using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that checks for existing PivotTables in a worksheet and removes the one at a given index using Aspose.Cells. | Show how to call worksheet.PivotTables.RemoveAt to delete a particular PivotTable and then save the workbook. | Provide a robust example that validates the PivotTable collection before invoking RemoveAt to avoid runtime errors.
// Common Searches: asp.net aspose.cells how to remove a pivot table at index 0 | c# code sample for deleting a pivot table from an Excel file using Aspose.Cells | remove specific pivot table from worksheet without affecting other data aspose.cells | check pivot table count before calling RemoveAt in Aspose.Cells C# | save changes after deleting pivot table with Aspose.Cells
// Tags: Aspose.Cells PivotTables.RemoveAt method | C# delete Excel pivot table programmatically | delete pivot table at given index | validate pivot table count before removal | save workbook after pivot table deletion

using Aspose.Cells;

// The example loads 'input.xlsx', verifies that the first worksheet contains at least one PivotTable, removes the PivotTable at index 0 using worksheet.PivotTables.RemoveAt, and saves the updated workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the target worksheet (e.g., the first one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete the PivotTable at the desired index (e.g., index 0)
        if (worksheet.PivotTables.Count > 0)
        {
            worksheet.PivotTables.RemoveAt(0);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
