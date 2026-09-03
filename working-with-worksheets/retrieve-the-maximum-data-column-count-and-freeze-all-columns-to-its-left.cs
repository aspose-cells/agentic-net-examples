// Title: Freeze columns left of the last used column in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, determines the last column containing data, and freezes all columns to its left using FreezePanes. | Create a C# example that calculates worksheet.MaxDataColumn, converts it to a column count, and applies FreezePanes to lock those columns in an existing workbook.
// Common Searches: Aspose.Cells C# freeze panes up to the last data column | How to programmatically freeze columns before the max data column in Excel using Aspose.Cells | C# get maximum data column index in a worksheet with Aspose.Cells | Freeze first N columns in Excel file with Aspose.Cells .NET API | Determine last used column and apply FreezePanes in Aspose.Cells C#
// Tags: Aspose.Cells column count detection | C# freeze left columns Excel Aspose | Worksheet dynamic column freezing | Excel final column detection Aspose.Cells .NET | programmatic column freeze Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, uses worksheet.MaxDataColumn to find the number of columns that contain data, and then calls FreezePanes to lock all columns to the left of that column. It ensures the output directory exists, saves the modified workbook, and includes basic error handling for missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
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
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the maximum column index that contains data (0‑based) and convert to count
                int maxDataColumnCount = worksheet.Cells.MaxDataColumn + 1;

                // Freeze all columns to the left of the maximum data column
                // FreezePanes(row, column, totalRows, totalColumns)
                // row = 0 (no row freeze), column = maxDataColumnCount (first unfrozen column)
                // totalRows = 0, totalColumns = maxDataColumnCount (number of columns to freeze)
                worksheet.FreezePanes(0, maxDataColumnCount, 0, maxDataColumnCount);

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
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
