// Title: Delete blank rows and columns on the first worksheet with DeleteOptions (UpdateReference) using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds data with empty rows and columns, then removes those empty rows and columns using the appropriate Aspose.Cells methods with an options object that updates references. | Show how to keep formulas intact while cleaning up blank rows and columns in Aspose.Cells. | Provide a complete example that saves the cleaned workbook as an XLSX file after removing all blank rows and columns.
// Common Searches: Aspose.Cells C# delete empty rows and columns while keeping formula references | How to use DeleteOptions.UpdateReference with DeleteBlankRows in .NET | Remove blank columns from the initial worksheet using Aspose.Cells API | Sample code for deleting blank rows and columns in a workbook with Aspose.Cells for .NET
// Tags: Aspose.Cells DeleteBlankRows DeleteOptions | Aspose.Cells DeleteBlankColumns UpdateReference | remove empty rows C# Aspose.Cells | clean worksheet blank columns .NET | delete blank rows columns XLSX Aspose

using System;
using Aspose.Cells;

// The example creates a new workbook, inserts sample data with intentional empty rows and columns, configures DeleteOptions to update references, removes all blank rows and columns from the first worksheet, and saves the result as DeletedBlankRowsColumns.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with intentional blank rows and columns
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("Data1");
        // Row 3 is left blank
        cells["A4"].PutValue("Data2"); // Row 4 contains data
        // Column B is left blank
        cells["C1"].PutValue("ColCHeader");
        cells["C2"].PutValue("ColCData");

        // Configure DeleteOptions (e.g., update references after deletion)
        DeleteOptions options = new DeleteOptions
        {
            UpdateReference = true
        };

        // Delete blank rows and blank columns using the configured options
        cells.DeleteBlankRows(options);
        cells.DeleteBlankColumns(options);

        // Save the modified workbook
        workbook.Save("DeletedBlankRowsColumns.xlsx", SaveFormat.Xlsx);
    }
}
