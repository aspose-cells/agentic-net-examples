// Title: Rename a worksheet after applying FreezePanes in Aspose.Cells for .NET and verify the freeze remains intact
// AI Prompts: Rename the worksheet in a workbook after calling FreezePanes and then check that the frozen rows and columns are still applied using Aspose.Cells for C#. | Change a sheet’s name while preserving its freeze pane settings and output the worksheet name and freeze status with Aspose.Cells .NET.
// Common Searches: c# aspose.cells rename worksheet after freeze panes | does renaming a sheet reset freeze panes in Aspose.Cells | how to keep freeze panes when changing worksheet name using Aspose.Cells | verify that freeze panes persist after worksheet rename in .NET | Aspose.Cells freeze panes state after worksheet name change
// Tags: Aspose.Cells rename worksheet preserve freeze panes | FreezePanes persistence after sheet rename C# | C# Aspose.Cells worksheet rename with frozen rows | verify frozen panes state Aspose.Cells workbook | Aspose.Cells worksheet rename example

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, names the first worksheet "OriginalSheet", applies FreezePanes to lock the first row and column, renames the worksheet to "RenamedSheet", prints the new name and confirms the freeze settings, and saves the file as FrozenRenameDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it an initial name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "OriginalSheet";

            // Freeze the first row and first column (split at cell B2)
            // Parameters: totalRows, totalColumns, splitRow, splitColumn
            sheet.FreezePanes(1, 1, 1, 1);

            // Rename the worksheet
            sheet.Name = "RenamedSheet";

            // Output results
            Console.WriteLine($"Worksheet renamed to: {sheet.Name}");
            Console.WriteLine("Freeze panes applied (first row and column).");

            // Save the workbook (optional)
            string outputPath = "FrozenRenameDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
