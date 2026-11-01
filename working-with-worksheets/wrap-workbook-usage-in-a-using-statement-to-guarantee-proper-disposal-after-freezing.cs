// Title: Use a using block to automatically dispose an Aspose.Cells Workbook after freezing the first row in C#
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook, applies FreezePanes to the first row, saves the file, and wraps the workbook in a using statement for automatic disposal. | Provide a concise example demonstrating how a using block can manage the lifecycle of an Aspose.Cells Workbook while freezing panes in .NET.
// Common Searches: C# Aspose.Cells apply FreezePanes to the top row and manage workbook lifecycle | using block for workbook disposal after FreezePanes in Aspose.Cells | best practice for handling Aspose.Cells workbook resources in .NET
// Tags: using pattern Aspose.Cells workbook cleanup | freeze panes first row Aspose.Cells C# | auto dispose workbook after FreezePanes | save workbook with frozen rows Aspose.Cells

using System;
using Aspose.Cells; // Aspose.Cells namespace

// // Demonstrates creating a Workbook, freezing the first row, saving to "FrozenWorkbook.xlsx", and guaranteeing disposal via a using block.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and ensure it is disposed properly
            using (Workbook workbook = new Workbook())
            {
                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Example: put some data into a cell
                sheet.Cells["A1"].PutValue("Sample Data");

                // Freeze the first row (row index 1, column index 0)
                // totalRows = 1 (freeze one row), totalColumns = 0 (no column freeze)
                sheet.FreezePanes(1, 0, 1, 0);

                // Save the workbook to a file
                workbook.Save("FrozenWorkbook.xlsx");
            } // Workbook is automatically disposed here after freezing and saving.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
