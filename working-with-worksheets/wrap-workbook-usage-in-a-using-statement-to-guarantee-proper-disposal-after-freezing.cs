// Title: Use a C# using statement to automatically dispose an Aspose.Cells Workbook after freezing the first row and column
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook inside a using construct, applies FreezePanes to the top‑left cell, and saves the file. | Refactor existing Aspose.Cells workbook code so the Workbook object is wrapped in a using clause for automatic disposal after freezing panes.
// Common Searches: how to ensure Aspose.Cells workbook is disposed after applying FreezePanes in C# | C# using statement for Aspose.Cells workbook cleanup after freezing panes | best practice for disposing Aspose.Cells Workbook when freezing first row and column
// Tags: Aspose.Cells using statement for workbook disposal | FreezePanes API Aspose.Cells C# | automatic workbook cleanup with using block | C# resource management Aspose.Cells

using Aspose.Cells;

// // Demonstrates creating a Workbook inside a using block, freezing the first row and column via FreezePanes, saving to "output.xlsx", and relying on automatic disposal of the Workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook inside a using block to guarantee disposal after operations
        using (var workbook = new Workbook())
        {
            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Freeze the first row and first column (freeze pane at cell A2)
            sheet.FreezePanes(1, 1, 1, 1);

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        } // Workbook is automatically disposed here
    }
}
