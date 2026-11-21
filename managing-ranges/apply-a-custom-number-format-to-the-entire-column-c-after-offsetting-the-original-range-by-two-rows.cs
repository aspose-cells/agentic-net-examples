// Title: Apply a custom number format to column C starting from row 3 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply custom number format to column C from row 3 | offset a range by two rows and set number format in Aspose.Cells | create dynamic range to last used row and apply style Aspose.Cells .NET | set thousand separator format for an entire column in an existing workbook using Aspose.Cells | apply style with custom number format to a column after skipping header rows in C#
// Tags: custom number format Aspose.Cells .NET | apply style to column range Aspose.Cells | offset range by rows Aspose.Cells | dynamic range based on MaxDataRow Aspose.Cells | format column C with thousand separator Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, defines a style with the custom number format "#,##0.00", builds a range that starts at row 3 column C and extends to the last used row, applies the style to that range, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure there is at least one worksheet
            Worksheet sheet;
            if (workbook.Worksheets.Count > 0)
            {
                sheet = workbook.Worksheets[0];
            }
            else
            {
                // Add a new worksheet and retrieve it by its index
                int newIndex = workbook.Worksheets.Add();
                sheet = workbook.Worksheets[newIndex];
            }

            // Create a custom number format style (e.g., thousand separator with two decimals)
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0.00";

            // Apply only the number format part of the style
            StyleFlag flag = new StyleFlag { NumberFormat = true };

            // Define the range start (row 3, column C) – zero‑based indexes
            int startRow = 2;      // third row in Excel
            int startColumn = 2;   // column C

            // Calculate how many rows to include from startRow to the last used row
            int totalRows = sheet.Cells.MaxDataRow - startRow + 1;
            if (totalRows < 0) totalRows = 0; // handle empty sheets

            // Create the range (use fully qualified Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range range = sheet.Cells.CreateRange(startRow, startColumn, totalRows, 1);

            // Apply the custom number format to the range
            range.ApplyStyle(customStyle, flag);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
