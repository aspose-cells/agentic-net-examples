// Title: Dynamically freeze columns up to the last populated column in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file, uses Worksheet.Cells.MaxDataColumn to locate the final column containing data, and applies FreezePanes to lock all columns through that index. | Create a reusable C# method that takes input and output paths, validates the source file, determines the maximum data column with Aspose.Cells, freezes the appropriate columns, and saves the workbook. | Show how to handle missing input files and ensure the output directory exists while dynamically applying FreezePanes based on the detected data column in Aspose.Cells.
// Common Searches: asp.net freeze panes based on last used column Aspose.Cells | c# get last populated column in Excel worksheet using Aspose.Cells | how to programmatically freeze columns up to max data column with Aspose.Cells | example of Worksheet.FreezePanes with dynamic column range in .NET
// Tags: dynamic column freeze Aspose.Cells | detect last populated column Aspose.Cells | freeze panes up to last column .NET | excel workbook manipulation C# Aspose | auto freeze columns based on content

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, retrieves the zero‑based index of the last column that contains data via Worksheet.Cells.MaxDataColumn, and freezes all columns up to that column using FreezePanes. It also creates the output directory if needed, saves the modified workbook to output.xlsx, and logs any errors.
class DynamicFreezeColumns
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or modify as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the last column that contains data (zero‑based index)
            int maxDataColumn = worksheet.Cells.MaxDataColumn;

            // Freeze all columns that contain data.
            // Use the overload with four parameters: row, column, totalRows, totalColumns.
            // Setting row = 0 freezes no rows; column = maxDataColumn + 1 freezes up to that column.
            worksheet.FreezePanes(0, maxDataColumn + 1, 0, 0);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
