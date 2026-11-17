// Title: C# example to verify Excel cell tooltips appear only when text is truncated using Aspose.Cells
// Common Searches: how to validate that a cell comment tooltip is shown only for truncated text in Aspose.Cells C# | detect overflow of cell text based on column width and check comment visibility with Aspose.Cells | Aspose.Cells C# code to iterate used range and warn if tooltip visibility doesn't match text truncation | C# program to compare cell string length to column width and ensure tooltip appears correctly | verify Excel tooltip behavior for truncated cells using Aspose.Cells library
// Tags: cell text truncation detection Aspose.Cells | comment tooltip visibility validation C# | column width based overflow check Excel | iterate used range Aspose.Cells workbook | excel tooltip verification example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, walks through the used range, determines if a cell's displayed text exceeds its column width, and checks whether the cell's comment (tooltip) is visible, reporting any mismatches between truncation and tooltip presence.
class TooltipVerification
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"[Error] Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Error] Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the used range
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through each cell in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                try
                {
                    Cell cell = cells[row, col];

                    // Skip empty cells
                    if (cell.Type == CellValueType.IsNull) continue;

                    // Check if the cell's displayed text is truncated
                    bool isTruncated = IsCellTextTruncated(sheet, cell);

                    // Retrieve the tooltip flag from the cell's comment visibility
                    Comment comment = sheet.Comments[row, col];
                    bool hasTooltip = comment != null && comment.IsVisible;

                    // Verify that tooltip appears only when text is truncated
                    if (isTruncated && !hasTooltip)
                    {
                        Console.WriteLine($"[Warning] Cell {cell.Name} is truncated but tooltip is missing.");
                    }
                    else if (!isTruncated && hasTooltip)
                    {
                        Console.WriteLine($"[Warning] Cell {cell.Name} is not truncated but tooltip is present.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[Error] Processing cell at row {row}, column {col}: {ex.Message}");
                }
            }
        }
    }

    // Determines whether the cell's displayed text would be truncated based on column width.
    static bool IsCellTextTruncated(Worksheet sheet, Cell cell)
    {
        // Get the text that Excel would display for the cell
        string displayedText = cell.StringValue ?? string.Empty;

        // Obtain the column width in character units (approximation)
        double columnWidth = sheet.Cells.GetColumnWidth(cell.Column);

        // If the text length exceeds the column width, it will be truncated.
        return displayedText.Length > columnWidth;
    }
}
