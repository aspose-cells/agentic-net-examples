// Title: Rename an Excel worksheet after setting freeze panes while preserving the frozen view using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# rename worksheet without losing freeze pane settings | how to keep freeze panes after changing worksheet name in Aspose.Cells | C# example to apply FreezePanes then rename sheet in Aspose.Cells | preserve frozen rows and columns when renaming a worksheet with Aspose.Cells | Aspose.Cells workbook freeze panes persist after worksheet rename
// Tags: freeze panes worksheet rename Aspose.Cells | frozen panes retention after worksheet rename | Aspose.Cells maintain freeze panes on sheet rename | C# Excel freeze panes persistence | save workbook with frozen panes Aspose.Cells

using System;
using Aspose.Cells;

// This example creates a new Workbook, applies FreezePanes to the first row and column, renames the worksheet, and saves the file, demonstrating that the freeze pane settings remain intact after the worksheet name change.
class WorksheetRenameAndFreezeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first row and first column (freeze panes at row 2, column 2)
            sheet.FreezePanes(1, 1, 1, 1);

            // Rename the worksheet
            sheet.Name = "RenamedSheet";

            // Output results
            Console.WriteLine("Worksheet renamed to: " + sheet.Name);
            Console.WriteLine("Freeze panes applied.");

            // Save the workbook
            string outputPath = "RenamedAndFrozen.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
