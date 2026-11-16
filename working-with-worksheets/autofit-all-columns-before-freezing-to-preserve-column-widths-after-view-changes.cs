// Title: Auto‑fit all columns then freeze the first row and column in an Excel file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells auto fit columns before applying FreezePanes | How to preserve column widths when freezing panes with Aspose.Cells .NET | AutoFitColumns followed by FreezePanes example in C# | Set column widths then freeze first row and column using Aspose.Cells | Aspose.Cells keep column sizing after FreezePanes in C#
// Tags: auto-fit columns Aspose.Cells C# | freeze panes preserve column width Aspose.Cells | Aspose.Cells AutoFitColumns method | Aspose.Cells FreezePanes usage | Excel worksheet column sizing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, auto‑fits all columns on the first worksheet to retain their widths, freezes the first row and first column (making B2 the top‑left scrollable cell), and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or iterate through all worksheets as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Auto‑fit all columns in the worksheet to preserve column widths
            sheet.AutoFitColumns();

            // Freeze the first row and first column.
            // row = 1, column = 1 => cell B2 is the top‑left scrollable cell.
            // totalRows = 1, totalColumns = 1 => freeze one row and one column.
            sheet.FreezePanes(1, 1, 1, 1);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
