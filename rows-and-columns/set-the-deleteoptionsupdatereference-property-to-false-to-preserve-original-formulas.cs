// Title: Aspose.Cells .NET – Keep Formulas Intact When Deleting Columns Using DeleteOptions.UpdateReference = false
// Description: Demonstrates how to create a workbook, add values and a formula, configure DeleteOptions with UpdateReference set to false, and delete a column without altering the original formula reference. The resulting file shows the formula unchanged after the column removal.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference false | preserve formulas .NET | delete column without adjusting formulas | C# Aspose.Cells example | formula reference stability | Excel automation DeleteOptions | Aspose.Cells row deletion | keep cell references after delete | Aspose.Cells API DeleteOptions
// Common Searches: Aspose.Cells keep formula after deleting column | DeleteOptions.UpdateReference false C# example | prevent formula shift when removing rows Aspose.Cells | how to disable reference update in Aspose.Cells | Aspose.Cells delete column without changing formulas
// Developer Intent: Configure DeleteOptions.UpdateReference = false so that row or column deletions do not modify existing cell formulas.
// Use Cases: Maintain legacy calculation links after removing a data column in a financial model. | Delete placeholder rows in a template while preserving dependent summary formulas. | Strip temporary helper columns from a report without breaking chart data sources.
// AI Prompts: Generate a C# snippet that deletes multiple rows while keeping all formulas unchanged using DeleteOptions.UpdateReference = false. | Explain how DeleteOptions.UpdateReference interacts with named ranges and merged cells in Aspose.Cells. | Show how to toggle DeleteOptions.UpdateReference based on a runtime condition before deleting a column.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add values and a formula, configure DeleteOptions with UpdateReference set to false, and delete a column without altering the original formula reference. The resulting file shows the formula unchanged after the column removal.
public class DeleteOptionsUpdateReferenceDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some data and a formula that references the data
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(20);
        cells["C1"].Formula = "=A1+B1"; // Formula will reference A1 and B1

        // Create DeleteOptions and set UpdateReference to false
        // This ensures that when we delete a column/row, existing formulas are NOT adjusted
        DeleteOptions deleteOptions = new DeleteOptions
        {
            UpdateReference = false
        };

        // Delete the first column (index 0) using the options above
        // After deletion, the formula in C1 will still be "=A1+B1" (referring to the original cells)
        worksheet.Cells.DeleteColumns(0, 1, deleteOptions);

        // Save the workbook to verify the result
        workbook.Save("DeleteOptionsFalseDemo.xlsx");
    }
}
