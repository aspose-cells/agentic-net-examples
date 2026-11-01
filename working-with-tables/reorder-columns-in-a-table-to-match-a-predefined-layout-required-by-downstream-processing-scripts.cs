// Title: Reorder Excel worksheet columns by header names using Aspose.Cells for .NET
// Common Searches: asp.net reorder Excel columns by header using Aspose.Cells | c# move Excel column to specific position based on header text Aspose.Cells | how to change column order in an existing .xlsx file with Aspose.Cells C# | programmatically reorder table columns to match predefined sequence Aspose.Cells | Aspose.Cells C# reorder columns without losing data
// Tags: reorder columns Aspose.Cells .xlsx | move column by header Aspose.Cells C# | insert column Aspose.Cells API | delete column Aspose.Cells API | find column index by header Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx workbook, locates columns by their header names (ID, Name, Date, Amount) on the first row, and reorders them to the desired sequence by copying, inserting, and deleting columns using Aspose.Cells for .NET, then saves the result as output.xlsx.
class ReorderTableColumns
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired column order by header names
            string[] desiredOrder = { "ID", "Name", "Date", "Amount" };

            // Header row index (0‑based)
            int headerRowIndex = 0;

            int targetPosition = 0;
            foreach (string header in desiredOrder)
            {
                int currentColumn = FindColumnIndex(worksheet, headerRowIndex, header);
                if (currentColumn == -1)
                {
                    // Header not found; skip
                    continue;
                }

                if (currentColumn != targetPosition)
                {
                    // Move the column manually (copy → insert → delete)
                    MoveColumn(worksheet, currentColumn, targetPosition);
                }

                // After moving, the column is at targetPosition
                targetPosition++;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Locate a column index by header text in a specific row
    static int FindColumnIndex(Worksheet sheet, int headerRow, string headerText)
    {
        int maxColumn = sheet.Cells.MaxColumn;
        for (int col = 0; col <= maxColumn; col++)
        {
            Cell cell = sheet.Cells[headerRow, col];
            if (cell != null && cell.Type == CellValueType.IsString &&
                cell.StringValue.Equals(headerText, StringComparison.OrdinalIgnoreCase))
            {
                return col;
            }
        }
        return -1; // Not found
    }

    // Manually move a column from sourceIndex to destinationIndex
    static void MoveColumn(Worksheet sheet, int sourceIndex, int destinationIndex)
    {
        // Determine the number of rows that contain data
        int maxRow = sheet.Cells.MaxDataRow;

        // Store source column values
        object[] columnData = new object[maxRow + 1];
        for (int row = 0; row <= maxRow; row++)
        {
            Cell srcCell = sheet.Cells[row, sourceIndex];
            columnData[row] = srcCell?.Value;
        }

        // Insert a new column at the destination index
        sheet.Cells.InsertColumn(destinationIndex);

        // Paste stored values into the new column
        for (int row = 0; row <= maxRow; row++)
        {
            sheet.Cells[row, destinationIndex].PutValue(columnData[row]);
        }

        // Adjust the index of the original column after insertion
        int deleteIndex = sourceIndex > destinationIndex ? sourceIndex + 1 : sourceIndex;

        // Delete the original column
        sheet.Cells.DeleteColumn(deleteIndex);
    }
}
