// Title: Delete blank columns on the first worksheet while updating cell references using DeleteOptions.UpdateReference in Aspose.Cells for .NET
// AI Prompts: Generate C# code that removes all empty columns from the first worksheet and automatically updates any formula references using DeleteOptions.UpdateReference. | Show how to call Cells.DeleteBlankColumns with a DeleteOptions object to preserve formulas after deleting blank columns in Aspose.Cells.
// Common Searches: Aspose.Cells C# delete empty columns and keep formulas updated | How to use DeleteOptions.UpdateReference when removing blank columns in a workbook | DeleteBlankColumns method example for first worksheet in .NET | Remove blank columns from worksheet without breaking cell references Aspose.Cells
// Tags: Aspose.Cells DeleteBlankColumns with UpdateReference | C# delete empty columns preserving formulas | DeleteOptions.UpdateReference usage in Aspose.Cells | remove blank columns first worksheet Aspose.Cells | Aspose.Cells workbook column cleanup .NET

using System;
using Aspose.Cells;

namespace AsposeCellsDeleteBlankColumnsExample
{
    // The example creates a workbook, places data in columns A and C leaving column B empty, configures DeleteOptions with UpdateReference set to true, deletes all blank columns on the first worksheet using Cells.DeleteBlankColumns, and saves the result as DeletedBlankColumns.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate data with a blank column (column B will be blank)
            cells["A1"].PutValue("Column A");
            cells["C1"].PutValue("Column C");
            cells["A2"].PutValue(1);
            cells["C2"].PutValue(3);

            // Set up DeleteOptions with UpdateReference = true
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete all blank columns on the first worksheet
            cells.DeleteBlankColumns(options);

            // Save the modified workbook
            workbook.Save("DeletedBlankColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}
