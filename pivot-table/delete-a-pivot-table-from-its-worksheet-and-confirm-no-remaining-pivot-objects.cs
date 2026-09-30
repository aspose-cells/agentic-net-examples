// Title: Delete a specific pivot table from an Excel worksheet using Aspose.Cells for .NET and verify that no pivot tables remain
// AI Prompts: Remove a pivot table at a specified index from a worksheet with Aspose.Cells in C# and save the updated workbook. | After deleting the pivot table, query the worksheet's PivotTables collection to confirm it is empty and output the result.
// Common Searches: Aspose.Cells C# delete pivot table by index example | How to remove a pivot table from an Excel file using Aspose.Cells .NET | Check for remaining pivot tables after deletion with Aspose.Cells | Save workbook after removing pivot tables in C# | Verify pivot table count in a worksheet using Aspose.Cells
// Tags: remove pivot table Aspose.Cells C# | pivot table deletion verification .NET | worksheet pivot tables count Aspose.Cells | save workbook after pivot table removal C# | delete pivot table by index Aspose.Cells | check empty pivot tables collection .NET

using System;
using Aspose.Cells;

// The example loads an Excel workbook, accesses a worksheet, deletes the first pivot table (or a specified one) using Aspose.Cells, confirms that the worksheet's PivotTables collection is empty, prints a status message, and saves the modified file.
class DeletePivotTableExample
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the pivot table (by index or name)
        Worksheet worksheet = workbook.Worksheets[0]; // or workbook.Worksheets["Sheet1"]

        // Ensure there is at least one pivot table before attempting deletion
        if (worksheet.PivotTables.Count > 0)
        {
            // Delete the first pivot table (or specify the index you want to remove)
            worksheet.PivotTables.RemoveAt(0);
        }

        // Verify that no pivot tables remain in the worksheet
        if (worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("Pivot table successfully deleted. No remaining pivot tables.");
        }
        else
        {
            Console.WriteLine($"There are still {worksheet.PivotTables.Count} pivot table(s) present.");
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
