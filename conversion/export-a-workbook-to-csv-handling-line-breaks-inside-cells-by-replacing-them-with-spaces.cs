// Title: Replace line breaks in Excel cells and export the workbook to CSV with Aspose.Cells for .NET
// AI Prompts: Iterate all worksheets, replace '\r' and '\n' in string cells with a space, then save the workbook as CSV using Aspose.Cells. | Clean newline characters from Excel cell values before performing a CSV export in a C# application with Aspose.Cells.
// Common Searches: Aspose.Cells .NET remove newline characters from Excel cells before CSV conversion | How to clean line breaks in Excel data when exporting to CSV using Aspose.Cells | C# replace line breaks in worksheet cells and save as CSV with Aspose.Cells library
// Tags: replace line breaks in Excel cells Aspose.Cells | CSV export with cleaned cell values .NET | iterate worksheets modify string cells Aspose | remove newline characters during CSV conversion Aspose.Cells | handle embedded line breaks in CSV output .NET

using Aspose.Cells;
using System;

// The example loads an Excel workbook, iterates through each worksheet and its string cells, replaces any '\r' or '\n' line‑break characters with a space, updates the cells, and finally saves the modified workbook as a CSV file using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection for the current sheet
            Cells cells = sheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Loop through all cells in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string cells (skip numbers, dates, etc.)
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;

                        // If the cell contains line breaks, replace them with spaces
                        if (!string.IsNullOrEmpty(text) && (text.Contains("\n") || text.Contains("\r")))
                        {
                            string cleaned = text
                                .Replace("\r\n", " ")
                                .Replace("\n", " ")
                                .Replace("\r", " ");

                            // Update the cell with the cleaned text
                            cell.PutValue(cleaned);
                        }
                    }
                }
            }
        }

        // Export the modified workbook to CSV format
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
