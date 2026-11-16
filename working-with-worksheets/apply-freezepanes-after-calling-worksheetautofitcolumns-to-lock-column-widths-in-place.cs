// Title: How to freeze the first row and column after auto‑fitting columns with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells freeze panes after AutoFitColumns example | lock column widths after auto‑fit using FreezePanes Aspose.Cells | determine used range for FreezePanes in Aspose.Cells C# | apply FreezePanes to first row and column after auto‑fit in .xlsx | Aspose.Cells C# save workbook with frozen panes and auto‑sized columns
// Tags: Aspose.Cells FreezePanes with used range | auto‑fit columns then freeze first row column C# | MaxDataRow MaxDataColumn for FreezePanes Aspose.Cells | save workbook with frozen panes Aspose.Cells | C# worksheet column width lock Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a Workbook, adds sample data, auto‑fits all columns, calculates the used row and column count via MaxDataRow/MaxDataColumn, and then calls FreezePanes(totalRows, totalColumns, 1, 1) to lock the first row and column. It ensures the output directory exists, saves the file as Output.xlsx, and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (optional)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");

            // Auto‑fit all columns to the content
            sheet.AutoFitColumns();

            // Freeze the first row and first column (split at cell B2)
            // totalRows and totalColumns can be set to the maximum used rows/columns
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(totalRows, totalColumns, 1, 1);

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
