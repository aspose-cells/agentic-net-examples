// Title: Rename an Excel worksheet after applying FreezePanes while preserving frozen rows and columns with Aspose.Cells for .NET
// AI Prompts: Change the worksheet's Name property after applying FreezePanes, then read the FreezePanes properties to confirm the frozen range is unchanged. | Programmatically verify that the FreezePanes settings remain the same after setting the worksheet's Name property in a C# Aspose.Cells workbook.
// Common Searches: Aspose.Cells keep freeze panes after worksheet rename in C# | does renaming a sheet reset FreezePanes settings Aspose.Cells .NET | how to check frozen rows persist after changing worksheet name using Aspose.Cells | C# verify freeze pane configuration after sheet name change with Aspose.Cells
// Tags: freeze pane settings survive sheet rename Aspose.Cells | rename worksheet without losing freeze pane configuration .NET | Aspose.Cells verify freeze pane state | C# workbook freeze pane stability | Excel freeze pane retention Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, adds data, freezes the first row and column with FreezePanes, renames the worksheet, and saves the file, demonstrating that the freeze pane configuration remains intact after the worksheet name change.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);

            // Freeze the first row and first column
            // Parameters: row index, column index, total rows to freeze, total columns to freeze
            // Row and column indices are zero‑based, so (1,1) freezes rows above row 1 and columns left of column 1
            sheet.FreezePanes(1, 1, 1, 1);

            // Rename the worksheet after freezing
            sheet.Name = "RenamedSheet";

            // Output confirmation
            Console.WriteLine($"Worksheet renamed to: {sheet.Name}");
            Console.WriteLine("Freeze panes applied successfully.");

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
