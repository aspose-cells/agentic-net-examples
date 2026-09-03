// Title: Delete columns left of a target column and then freeze the first column using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, remove columns 0 through N‑1, then apply FreezePanes at row 0, column 1 with Aspose.Cells in C#. | Write C# code that deletes all columns preceding column D and then freezes the first visible column using Aspose.Cells. | Create a script that trims left‑hand columns from a worksheet and sets a column‑only freeze pane without affecting rows, leveraging the Aspose.Cells FreezePanes overload.
// Common Searches: Aspose.Cells C# delete columns before a specific index and then set freeze pane on first column | How to remove left side columns and keep only column D visible while freezing column A with Aspose.Cells | C# example for deleting columns 0-2 and applying FreezePanes(0,1) using Aspose.Cells | Freeze pane after column deletion Aspose.Cells .NET tutorial
// Tags: delete columns before freeze pane Aspose.Cells C# | freeze panes column only Aspose.Cells | Aspose.Cells worksheet column removal | Aspose.Cells FreezePanes overload C# | Excel column deletion and freeze pane .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file, deletes all columns to the left of a specified index, freezes the first column with FreezePanes, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Define the column index where you want the freeze pane to start (0‑based)
            // Example: freeze at column D (index 3)
            int originalFreezeColumn = 3;

            // Delete all columns to the left of the desired freeze area
            // This removes columns 0 through originalFreezeColumn‑1
            // The third argument (true) ensures that only the column data is removed
            worksheet.Cells.DeleteColumns(0, originalFreezeColumn, true);

            // Freeze panes at the first row (no row freeze) and first column
            // Use the overload with four parameters (row, column, totalRows, totalColumns)
            // Setting totalRows and totalColumns to 0 freezes only the specified column.
            worksheet.FreezePanes(0, 1, 0, 0);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
