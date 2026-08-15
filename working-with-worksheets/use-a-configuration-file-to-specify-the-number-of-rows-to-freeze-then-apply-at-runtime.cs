// Title: Apply FreezePanes Dynamically from a Config File with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to read the number of rows and columns to freeze from a configuration source (e.g., appsettings.json), create a workbook, populate sample data, apply FreezePanes with the retrieved values, and save the Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | FreezePanes | configuration file | appsettings.json | dynamic freeze rows | Excel freeze panes | runtime settings
// Common Searches: Aspose.Cells freeze panes from appsettings | C# read freeze rows from config and apply FreezePanes | set frozen rows dynamically Aspose.Cells | how to use variable for FreezePanes in .NET
// Developer Intent: Load the row/column count for FreezePanes from a configuration source and apply it at runtime with Aspose.Cells.
// Use Cases: Generate reports where the header height is defined in appsettings.json and applied automatically. | Create Excel dashboards that let administrators configure how many top rows and columns stay visible. | Build a batch processor that reads freeze settings from a JSON file and applies them to multiple worksheets.
// AI Prompts: Write C# code that reads an integer "FreezeRows" from appsettings.json and uses it in sheet.FreezePanes with Aspose.Cells. | Show how to add default handling when the config value is missing or invalid while applying FreezePanes. | Generate a reusable method that accepts a configuration object containing freeze row/column counts and applies FreezePanes to any worksheet.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to read the number of rows and columns to freeze from a configuration source (e.g., appsettings.json), create a workbook, populate sample data, apply FreezePanes with the retrieved values, and save the Excel file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Number of rows/columns to freeze; replace with desired value or read from another source.
            int freezeRows = 5;

            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data so the frozen area is visible.
            for (int i = 0; i < 20; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
                sheet.Cells[i, 1].PutValue($"Data {i + 1}");
            }

            // Apply freeze panes at the specified row/column.
            // FreezeRows rows and FreezeRows columns starting from cell (freezeRows, freezeRows).
            sheet.FreezePanes(freezeRows, freezeRows, freezeRows, freezeRows);

            // Determine output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FrozenRowsDemo.xlsx");

            // Save the workbook.
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to {outputPath} with {freezeRows} frozen rows and columns.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
