// Title: C# – Use Aspose.Cells Workbook within a using block to auto‑dispose after FreezePanes
// Description: This example demonstrates creating an Aspose.Cells Workbook inside a C# using statement, adding data to the first worksheet, freezing panes at cell C3 (freezing two rows and two columns), saving the workbook as FrozenPaneDemo.xlsx, and letting the runtime automatically call Dispose to release unmanaged resources.
// Keywords: Aspose.Cells | C# | using statement | Workbook disposal | FreezePanes | resource cleanup | Excel automation | memory leak prevention | scoped workbook | Excel file generation
// Common Searches: Aspose.Cells using block C# | how to dispose Workbook after FreezePanes | C# auto‑dispose Aspose.Cells workbook | freeze panes with Aspose.Cells and ensure cleanup | resource management Aspose.Cells workbook | prevent memory leaks Aspose.Cells
// Developer Intent: Automatically release workbook resources after applying FreezePanes by wrapping the Workbook in a using block.
// Use Cases: Generate a single Excel file, apply FreezePanes, and guarantee cleanup with a using statement. | Process multiple workbooks in a batch, applying FreezePanes to each while preventing memory leaks. | Integrate workbook creation, data insertion, pane freezing, and saving into a scoped method for safe resource handling. | Wrap third‑party Excel operations in a deterministic disposal pattern for server‑side services.
// AI Prompts: Provide C# code that creates an Aspose.Cells Workbook, writes sample data, freezes panes at a specified cell, saves the file, and uses a using block for automatic disposal. | Show how to loop through a list of Excel files, open each with Aspose.Cells inside a using statement, apply FreezePanes, and save the changes. | Explain why wrapping an Aspose.Cells Workbook in a using statement is important for resource management after operations like FreezePanes. | Generate a reusable method that accepts a file path and freeze cell reference, creates a workbook, applies FreezePanes, saves, and ensures disposal.

using System;
using Aspose.Cells;

// This example demonstrates creating an Aspose.Cells Workbook inside a C# using statement, adding data to the first worksheet, freezing panes at cell C3 (freezing two rows and two columns), saving the workbook as FrozenPaneDemo.xlsx, and letting the runtime automatically call Dispose to release unmanaged resources.
class Program
{
    static void Main()
    {
        // Workbook is created inside a using block to ensure Dispose is called automatically
        using (Workbook workbook = new Workbook())
        {
            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Sample before freeze");
            sheet.Cells["C3"].PutValue("Freeze point");

            // Freeze panes at cell C3 (2 rows and 2 columns are frozen)
            sheet.FreezePanes("C3", 2, 2);

            // Save the workbook to a file
            workbook.Save("FrozenPaneDemo.xlsx");
        } // workbook.Dispose() is invoked here, releasing all unmanaged resources
    }
}
