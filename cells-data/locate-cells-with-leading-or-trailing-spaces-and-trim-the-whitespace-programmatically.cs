// Title: Programmatically trim leading and trailing spaces from string cells in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells remove whitespace from all string cells in a workbook | Trim leading and trailing spaces in Excel cells with Aspose.Cells API | Iterate through worksheets and clean string values using Aspose.Cells for .NET | Save cleaned Excel file after trimming cell text with Aspose.Cells | Handle missing input file when trimming cell strings in C#
// Tags: clean string cells Aspose.Cells | worksheet cell iteration C# Aspose | strip leading trailing spaces Excel .NET | update cell after trim Aspose | export cleaned workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (or creates a new workbook if the file is missing), iterates over every worksheet, row, and column, detects string‑type cells, trims leading and trailing whitespace using String.Trim(), writes the trimmed value back to the cell when changes are needed, and finally saves the cleaned workbook to a new file while handling load and save exceptions.
class TrimWhitespace
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            if (File.Exists(inputPath))
            {
                // Load the workbook from the existing file (no password required)
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook if the input file is missing
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Iterate through all worksheets and trim whitespace from string cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        string trimmed = original.Trim();
                        if (!original.Equals(trimmed))
                        {
                            cell.PutValue(trimmed);
                        }
                    }
                }
            }
        }

        try
        {
            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
