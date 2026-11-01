// Title: Automatically freeze the first 100 rows in Excel worksheets that contain more than 100 rows using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook, checks each worksheet’s row count, and applies FreezePanes to lock the top 100 rows when the sheet exceeds 100 rows. | Generate a method that iterates through all worksheets in an Aspose.Cells workbook and conditionally calls FreezePanes(100,0,…) for sheets with more than 100 data rows. | Create a script that validates the existence of an input Excel file, freezes the first hundred rows on large sheets, and saves the modified file using Aspose.Cells.
// Common Searches: Aspose.Cells C# freeze top 100 rows only if worksheet has more than 100 rows | How to apply conditional freeze panes in a .NET Excel workbook with Aspose | C# iterate all worksheets and set FreezePanes based on row count using Aspose.Cells | Freeze first hundred rows in Excel file programmatically with Aspose.Cells for .NET
// Tags: conditional freeze panes Aspose.Cells C# | freeze top rows Excel Aspose.Cells | apply FreezePanes based on row count | Aspose.Cells workbook row limit handling | C# iterate worksheets set freeze panes

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, checks each worksheet’s data rows, freezes the first 100 rows on sheets with more than 100 rows using FreezePanes, and saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the index of the last row that contains data (zero‑based)
                int lastDataRow = sheet.Cells.MaxDataRow;
                int totalRows = lastDataRow + 1; // Convert to 1‑based count

                // If the worksheet has more than 100 rows, apply freeze panes
                if (totalRows > 100)
                {
                    // Freeze the first 100 rows (rows 0‑99). No column freeze.
                    // Using the overload that requires row, column, totalRows, totalColumns.
                    sheet.FreezePanes(100, 0, 100, 0);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
