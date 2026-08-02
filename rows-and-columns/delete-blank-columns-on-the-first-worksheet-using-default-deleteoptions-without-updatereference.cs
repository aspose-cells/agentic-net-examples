// Title: C# – Delete Blank Columns on First Worksheet with Aspose.Cells (default DeleteOptions)
// Description: Creates a new workbook, writes data to columns A and C leaving column B empty, then calls Worksheet.Cells.DeleteBlankColumns() which uses the default DeleteOptions (UpdateReference = false) to remove the empty column, and saves the file as DeletedBlankColumns.xlsx.
// Keywords: Aspose.Cells | DeleteBlankColumns | C# | .NET | remove empty columns | default DeleteOptions | UpdateReference false | worksheet manipulation | Excel automation
// Common Searches: Aspose.Cells delete blank columns C# | How to remove empty columns without updating references in Aspose.Cells | Default behavior of DeleteBlankColumns method | Delete blank columns on first sheet Aspose.Cells .NET | Remove completely empty columns from Excel using Aspose.Cells
// Developer Intent: Delete all completely empty columns from the first worksheet while keeping existing cell references unchanged.
// Use Cases: Clean up a generated workbook by stripping out columns that contain no data before distribution. | Prepare a template where placeholder columns may be left empty and need automatic removal. | Produce compact reports by eliminating columns that were conditionally populated but ended up blank.
// AI Prompts: Write C# code that uses Aspose.Cells to delete blank columns on a specific worksheet without affecting formula references. | Explain how Worksheet.Cells.DeleteBlankColumns determines blank columns and what the default DeleteOptions settings imply. | Show how to configure DeleteOptions to keep cell references unchanged while deleting empty columns in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, writes data to columns A and C leaving column B empty, then calls Worksheet.Cells.DeleteBlankColumns() which uses the default DeleteOptions (UpdateReference = false) to remove the empty column, and saves the file as DeletedBlankColumns.xlsx.
    class DeleteBlankColumnsDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with data leaving column B blank
            worksheet.Cells["A1"].PutValue("Column A");
            worksheet.Cells["C1"].PutValue("Column C"); // Column B will be blank
            worksheet.Cells["A2"].PutValue("Data A");
            worksheet.Cells["C2"].PutValue("Data C");

            // Delete all blank columns using default DeleteOptions (UpdateReference = false)
            worksheet.Cells.DeleteBlankColumns();

            // Save the modified workbook
            workbook.Save("DeletedBlankColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}
