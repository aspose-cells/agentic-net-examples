// Title: Freeze Column A in a New Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Creates a Workbook, adds a header row and ten data rows, applies Worksheet.FreezePanes to lock the first column (A) while allowing horizontal scrolling, and saves the file as FreezeFirstColumn.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# freeze column | Worksheet.FreezePanes example | create workbook Aspose.Cells | lock first column Excel | save Excel file Aspose.Cells
// Common Searches: how to freeze first column with Aspose.Cells .NET | C# sample code for Worksheet.FreezePanes | freeze column A in generated Excel file | Aspose.Cells create workbook and lock column
// Developer Intent: Apply a freeze pane to the first column of a newly generated worksheet after populating it with sample data.
// Use Cases: Report where the ID column must stay visible while scrolling horizontally. | Export template that keeps reference data fixed for large tables. | Dashboard sheet with a persistent key column for quick lookup.
// AI Prompts: Generate C# code using Aspose.Cells to create a workbook, add headers and rows, and freeze column A. | Explain each parameter of Worksheet.FreezePanes and how they affect frozen rows and columns. | Extend the example to freeze both the first row and the first column and save the result.

using System;
using Aspose.Cells;

// Creates a Workbook, adds a header row and ten data rows, applies Worksheet.FreezePanes to lock the first column (A) while allowing horizontal scrolling, and saves the file as FreezeFirstColumn.xlsx using Aspose.Cells for .NET.
class FreezeFirstColumnDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample header row
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["C1"].PutValue("Score");

        // Add sample data rows
        for (int i = 2; i <= 10; i++)
        {
            worksheet.Cells[i - 1, 0].PutValue(i - 1);                     // ID
            worksheet.Cells[i - 1, 1].PutValue($"Item {i - 1}");          // Name
            worksheet.Cells[i - 1, 2].PutValue((i - 1) * 10);            // Score
        }

        // Freeze the first column (column A)
        // Parameters: row index, column index, frozen rows, frozen columns
        // Setting row index to 0 and column index to 1 freezes column A.
        worksheet.FreezePanes(0, 1, 0, 1);

        // Save the workbook to a file
        workbook.Save("FreezeFirstColumn.xlsx");
    }
}
