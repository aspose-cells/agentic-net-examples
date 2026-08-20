// Title: C# – Wrap Aspose.Cells Workbook in a using Block to Auto‑Dispose After Freezing Panes
// Description: This example shows how to create an Aspose.Cells Workbook inside a C# using statement, freeze panes at cell C3 (freezing two rows and two columns), optionally unfreeze them, save the file as FrozenDemo.xlsx, and let the runtime automatically call Dispose when the block ends.
// Keywords: Aspose.Cells using statement | C# workbook disposal | freeze panes Aspose.Cells | resource cleanup .NET | IDisposable Aspose.Cells | auto dispose workbook
// Common Searches: Aspose.Cells how to use using block | C# dispose workbook after freeze panes | auto‑release Aspose.Cells resources | freeze panes C3 Aspose.Cells example | best practice workbook cleanup Aspose
// Developer Intent: Ensure the Workbook object is automatically disposed after performing freeze‑pane operations without writing explicit Dispose calls.
// Use Cases: Freeze panes in a newly created workbook while guaranteeing resource release. | Perform multiple worksheet modifications (freeze, unfreeze, formatting) inside a using block for clean code. | Save a workbook after unfreezing panes and rely on the using statement for disposal.
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook, freezes panes at a specified cell, and uses a using block for automatic disposal. | Explain why Aspose.Cells Workbook implements IDisposable and how a using statement simplifies memory management. | Provide a sample that wraps several worksheet operations—including FreezePanes and UnFreezePanes—in a using block and saves the workbook.

using System;
using Aspose.Cells;

// This example shows how to create an Aspose.Cells Workbook inside a C# using statement, freeze panes at cell C3 (freezing two rows and two columns), optionally unfreeze them, save the file as FrozenDemo.xlsx, and let the runtime automatically call Dispose when the block ends.
class Program
{
    static void Main()
    {
        // Workbook is created inside a using block to ensure Dispose is called automatically
        using (Workbook workbook = new Workbook())
        {
            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze panes at cell C3 (rows 2, columns 2 are frozen)
            sheet.FreezePanes("C3", 2, 2);

            // Additional operations can be performed here

            // Unfreeze panes before saving (optional)
            sheet.UnFreezePanes();

            // Save the workbook to a file
            workbook.Save("FrozenDemo.xlsx");
        } // workbook.Dispose() is invoked here
    }
}
