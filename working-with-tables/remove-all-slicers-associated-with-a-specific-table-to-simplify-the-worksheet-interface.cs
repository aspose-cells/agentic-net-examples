// Title: Remove all slicers from a specific table or worksheet with Aspose.Cells for .NET
// Description: Loads an Excel workbook, selects the first worksheet, checks for tables, accesses the worksheet's SlicerCollection, clears every slicer, and saves the updated file. Use this pattern to delete slicers tied to a table or the entire sheet.
// Keywords: Aspose.Cells remove slicers | clear slicer collection .NET | delete Excel slicers programmatically | Aspose.Cells slicer Clear method | C# remove worksheet slicers
// Common Searches: how to delete slicers in Aspose.Cells C# | clear all slicers from an Excel sheet using Aspose.Cells | remove slicers linked to a table with Aspose.Cells .NET | Aspose.Cells example for removing slicers | C# code to clear slicer collection in Excel
// Developer Intent: Delete slicers associated with a table or the whole worksheet in an Excel file using Aspose.Cells.
// Use Cases: Prepare a workbook for publishing by stripping interactive slicers that are no longer needed. | Reset UI filters after a data refresh to avoid stale selections before exporting. | Automate workbook cleanup for PDF or CSV conversion where slicers would not appear.
// AI Prompts: Generate C# code with Aspose.Cells that removes only slicers belonging to a given ListObject instead of clearing the entire collection. | Show how to iterate through SlicerCollection and delete slicers whose TableName matches a specified table. | Explain how to verify the presence of slicers on a worksheet before invoking the Clear() method.

using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

// Loads an Excel workbook, selects the first worksheet, checks for tables, accesses the worksheet's SlicerCollection, clears every slicer, and saves the updated file. Use this pattern to delete slicers tied to a table or the entire sheet.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the target table
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one table
        if (worksheet.ListObjects.Count == 0)
        {
            // No tables – nothing to do
            workbook.Save("output.xlsx");
            return;
        }

        // Get the specific table (for example, the first one)
        ListObject targetTable = worksheet.ListObjects[0];

        // Obtain the slicer collection for the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Remove all slicers that belong to the worksheet.
        // If slicers are only associated with the target table, this effectively removes them.
        // Otherwise, you could iterate and remove conditionally.
        slicers.Clear();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
