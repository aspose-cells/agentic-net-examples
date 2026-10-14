// Title: Delete columns left of a target column and apply FreezePanes in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Delete all columns before column index 2, then invoke Worksheet.FreezePanes(1,1,1,1) with Aspose.Cells in C#. | Remove preceding columns, update the worksheet layout, and set freeze panes at the first row and column using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# delete columns before freezing panes example | How to remove left-side columns and then freeze top row in Excel with Aspose.Cells | C# code to delete columns and set FreezePanes using Aspose.Cells | Adjust column index after DeleteColumns before calling FreezePanes in Aspose.Cells
// Tags: Aspose.Cells DeleteColumns usage | Aspose.Cells FreezePanes overload | C# Excel column removal with Aspose.Cells | Adjust freeze pane position after column deletion | Aspose.Cells worksheet column index update

using Aspose.Cells;
using System;
using System.IO;

// The program loads an Excel workbook, deletes every column to the left of a specified freeze column, then applies FreezePanes at the first row and column using Aspose.Cells for .NET, and finally saves the modified file while handling missing input files and exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Desired freeze column (0‑based). Example: freeze at column C (index 2)
            int originalFreezeColumn = 2;

            // Delete all columns that are to the left of the freeze column
            if (originalFreezeColumn > 0)
            {
                // DeleteColumns(startColumn, totalColumns, updateReference)
                sheet.Cells.DeleteColumns(0, originalFreezeColumn, false);
            }

            // Freeze panes at the first row and first column
            // Using the overload that specifies split position and number of rows/columns to freeze
            sheet.FreezePanes(1, 1, 1, 1);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
