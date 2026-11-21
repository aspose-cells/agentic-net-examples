// Title: Freeze the first ten columns of worksheets that exceed ten columns using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells freeze first ten columns when worksheet has more than 10 columns | How to apply FreezePanes conditionally based on column count in Aspose.Cells | Iterate all worksheets and set freeze pane for wide sheets using Aspose.Cells .NET | Programmatically freeze columns in Excel with Aspose.Cells if column count > 10 | Aspose.Cells .NET freeze pane only on sheets with over ten used columns
// Tags: conditional freeze pane Aspose.Cells | detect max column Aspose.Cells | apply FreezePanes across worksheets | Excel column based freeze using .NET | Aspose.Cells worksheet iteration for freeze

using Aspose.Cells;
using System;
using System.IO;

// The C# program loads an input workbook, iterates each worksheet, determines the number of used columns, and if a sheet contains more than ten columns it freezes the first ten columns (no rows) using the FreezePanes method, then saves the modified workbook to an output file.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the number of columns that contain data (zero‑based index + 1)
                int columnCount = sheet.Cells.MaxColumn + 1;

                // If the worksheet has more than ten columns, apply a freeze pane
                if (columnCount > 10)
                {
                    // Freeze the first ten columns (no rows are frozen)
                    // FreezePanes(totalRows, totalColumns, row, column)
                    sheet.FreezePanes(0, 10, 0, 10);
                }
            }

            // Output workbook path
            string outputPath = "output.xlsx";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
