// Title: Format pivot table header row and column with red bold font and yellow background using Aspose.Cells for .NET
// Common Searches: how to change pivot table header font color and background in Aspose.Cells C# | apply custom style to first row and first column of a pivot table using Aspose.Cells .NET | loop through PivotTable.TableRange1 to style header cells in C# | set red bold text for Excel pivot field headers with Aspose.Cells API | format pivot table row and column headers programmatically in .NET workbook
// Tags: Aspose.Cells pivot table header styling | C# apply cell style to PivotTable headers | Excel pivot field header formatting Aspose | PivotTable.TableRange1 iteration for styling | custom red bold yellow background style Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotField classes

// The example loads an Excel workbook, accesses the first worksheet's pivot table, creates a style with red bold font and yellow background, and iterates over the pivot table's TableRange1 to apply the style to the header row and header column before saving the updated file.
class PivotHeaderFormatter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the first worksheet contains the pivot table
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Create a style for the header cells
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = Color.Red;          // Red font color
            headerStyle.Font.IsBold = true;              // Bold text
            headerStyle.ForegroundColor = Color.Yellow; // Yellow background
            headerStyle.Pattern = BackgroundType.Solid;

            // Get the overall area of the pivot table (use TableRange1)
            CellArea tableArea = pivotTable.TableRange1;

            // Apply style to the first row (column headers)
            for (int col = tableArea.StartColumn; col <= tableArea.EndColumn; col++)
            {
                Cell cell = sheet.Cells[tableArea.StartRow, col];
                cell.SetStyle(headerStyle);
            }

            // Apply style to the first column (row headers)
            for (int row = tableArea.StartRow; row <= tableArea.EndRow; row++)
            {
                Cell cell = sheet.Cells[row, tableArea.StartColumn];
                cell.SetStyle(headerStyle);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
