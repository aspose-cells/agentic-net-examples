// Title: How to delete a specific row in Aspose.Cells for .NET while ensuring DeleteOptions.UpdateReference is true
// AI Prompts: Generate C# code that creates a workbook, fills sample data, and deletes row 2 using Aspose.Cells DeleteRows with a DeleteOptions instance whose UpdateReference property is set to true. | Demonstrate how to inspect the DeleteOptions.UpdateReference flag before calling DeleteRows and raise an error if the flag is false. | Provide a complete example that removes rows, updates all formula references, and includes robust try‑catch error handling with Aspose.Cells.
// Common Searches: Aspose.Cells C# delete row and keep formulas updated using DeleteOptions | Verify DeleteOptions.UpdateReference before removing rows in a .NET workbook | How to use DeleteRows with the UpdateReference flag in Aspose.Cells for C# | Delete a row in an Excel file with Aspose.Cells while preserving cell references | C# Aspose.Cells DeleteRows example with DeleteOptions object
// Tags: Aspose.Cells DeleteRows DeleteOptions | C# DeleteOptions usage | preserve formulas after row deletion | row removal update references .NET | exception handling Aspose.Cells workbook modification

using Aspose.Cells;
using System;

// The example creates a workbook, populates cells A1‑B3, configures a DeleteOptions object with UpdateReference set to true, deletes the second row using DeleteRows, and saves the file as Result.xlsx, all wrapped in proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue(100);
            sheet.Cells["A2"].PutValue(200);
            sheet.Cells["A3"].PutValue(300);
            sheet.Cells["B1"].PutValue(1);
            sheet.Cells["B2"].PutValue(2);
            sheet.Cells["B3"].PutValue(3);

            // Delete the second row (index 1) and update references
            sheet.Cells.DeleteRows(1, 1, true);

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
