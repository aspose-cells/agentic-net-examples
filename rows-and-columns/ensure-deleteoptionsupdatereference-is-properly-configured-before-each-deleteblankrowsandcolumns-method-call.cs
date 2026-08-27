// Title: Configure DeleteBlankOptions.UpdateReference before deleting blank rows and columns using Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets DeleteBlankOptions.UpdateReference to true and then calls Cells.DeleteBlankRows and Cells.DeleteBlankColumns. | Show how to reuse a single DeleteBlankOptions object for both row and column blank removal while keeping formulas intact. | Describe the impact of EmptyStringAsBlank and UpdateReference settings on worksheet references after deleting blanks.
// Common Searches: Aspose.Cells how to keep formulas when deleting blank rows in C# | Set DeleteBlankOptions.UpdateReference property before DeleteBlankRows Aspose.Cells .NET | Reuse DeleteBlankOptions for deleting both rows and columns in a workbook | Effect of EmptyStringAsBlank option on blank row removal in Aspose.Cells | Update cell references after removing empty columns with Aspose.Cells
// Tags: DeleteBlankOptions UpdateReference configuration | delete blank rows and columns Aspose.Cells | preserve formulas during blank deletion | reuse DeleteBlankOptions instance for rows and columns | EmptyStringAsBlank effect Aspose.Cells | cell reference update after blank removal

using System;
using Aspose.Cells;

namespace DeleteBlankRowsColumnsDemo
{
    // // Demonstrates creating a workbook, adding data with blank rows and columns, configuring DeleteBlankOptions (UpdateReference=true, EmptyStringAsBlank=true), deleting blank rows and columns using the same options, and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data with blank rows and columns
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("");               // Blank row
            cells["A3"].PutValue("Data");
            cells["B1"].PutValue("");               // Blank column
            cells["C1"].PutValue("Another Header");
            cells["C2"].PutValue("More Data");

            // Configure DeleteBlankOptions with UpdateReference set to true
            DeleteBlankOptions deleteOptions = new DeleteBlankOptions
            {
                UpdateReference = true,          // Ensure references are updated
                EmptyStringAsBlank = true,       // Treat empty strings as blanks
                DrawingsAsBlank = true           // Default behavior
            };

            // Delete blank rows using the configured options
            cells.DeleteBlankRows(deleteOptions);

            // Delete blank columns using the same configured options
            cells.DeleteBlankColumns(deleteOptions);

            // Save the workbook to verify the result
            workbook.Save("DeletedBlankRowsAndColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}
