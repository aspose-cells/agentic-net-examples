// Title: Delete Blank Rows in First Worksheet with Aspose.Cells C# (default DeleteOptions)
// Description: Creates a workbook, adds data with intentional empty rows, then calls Cells.DeleteBlankRows() using the default DeleteOptions (UpdateReference = false) to remove only the blank rows and saves the file as DeletedBlankRows.xlsx.
// Keywords: Aspose.Cells | DeleteBlankRows | C# | .NET | blank rows removal | default DeleteOptions | UpdateReference false | worksheet cleanup
// Common Searches: Aspose.Cells delete blank rows C# | DeleteBlankRows default options | Remove empty rows without updating references Aspose.Cells | How to clean up worksheet rows in .NET
// Developer Intent: Remove all empty rows from the first worksheet while leaving existing cell references unchanged.
// Use Cases: Clean imported spreadsheets that contain sporadic empty rows before analysis. | Prepare data for charts or pivot tables where blank rows cause gaps. | Automate workbook sanitization in ETL pipelines to ensure consistent row structures.
// AI Prompts: Generate C# code using Aspose.Cells to delete blank rows on the first worksheet without altering cell references. | Show how to verify that only empty rows were removed after calling DeleteBlankRows. | Explain the effect of DeleteOptions.UpdateReference when deleting blank rows in Aspose.Cells.

using System;
using Aspose.Cells;

namespace DeleteBlankRowsExample
{
    // Creates a workbook, adds data with intentional empty rows, then calls Cells.DeleteBlankRows() using the default DeleteOptions (UpdateReference = false) to remove only the blank rows and saves the file as DeletedBlankRows.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data with intentional blank rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data1");
            // Row 3 is left blank
            cells["A4"].PutValue("Data2");
            // Row 5 is left blank
            cells["A6"].PutValue("Data3");

            // Delete all blank rows using the default DeleteOptions.
            // The default options have UpdateReference = false, which satisfies the requirement.
            cells.DeleteBlankRows();

            // Save the modified workbook
            workbook.Save("DeletedBlankRows.xlsx", SaveFormat.Xlsx);
        }
    }
}
