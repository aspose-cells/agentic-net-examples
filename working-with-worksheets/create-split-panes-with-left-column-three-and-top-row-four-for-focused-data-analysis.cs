// Title: Freeze the first four rows and first three columns (split panes) in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# freeze first four rows and first three columns | how to set split panes in an Excel file with Aspose.Cells .NET | C# example for freezing top rows and left columns using Aspose.Cells | using Aspose.Cells FreezePanes to create frozen panes at row 5 column 4 | explain parameters of Aspose.Cells FreezePanes method in C#
// Tags: Aspose.Cells FreezePanes API | split pane configuration .NET Excel | freeze top rows left columns Aspose.Cells | C# Excel worksheet frozen panes | Excel workbook split pane example

using System;
using Aspose.Cells;

namespace AsposeCellsSplitPaneExample
{
    // The sample creates a new Workbook, fills cells A1:J20 with sample data, then freezes the first four rows and first three columns by calling sheet.FreezePanes(4, 3, 0, 0), and saves the file as 'SplitPaneExample.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data for demonstration.
                for (int row = 0; row < 20; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Freeze the top 4 rows and the left 3 columns.
                // Use the overload with four parameters (row, column, totalRows, totalColumns).
                // Setting totalRows and totalColumns to 0 freezes rows above 'row' and columns left of 'column'.
                sheet.FreezePanes(4, 3, 0, 0);

                // Save the workbook to a file.
                workbook.Save("SplitPaneExample.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
