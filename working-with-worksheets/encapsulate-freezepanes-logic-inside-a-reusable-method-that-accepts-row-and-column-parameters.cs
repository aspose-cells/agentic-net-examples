// Title: How to build a reusable C# method that applies FreezePanes with custom row and column indexes in Aspose.Cells
// AI Prompts: Write a C# helper method that receives a Worksheet, a zero‑based row index, and a zero‑based column index, then calls Worksheet.FreezePanes to lock those rows and columns using Aspose.Cells. | Demonstrate calling the helper to freeze the first row and the first two columns of a newly created workbook and save the file as an .xlsx.
// Common Searches: asp.net aspose.cells how to freeze specific rows and columns with a reusable function | c# parameterized FreezePanes example using Aspose.Cells worksheet | create a utility method for Worksheet.FreezePanes overload in C# | freeze panes on first row and first two columns programmatically Aspose.Cells | reusable freeze panes code snippet for Aspose.Cells C# projects
// Tags: freeze panes helper Aspose.Cells | parameterized worksheet freeze rows columns | Worksheet.FreezePanes overload usage C# | custom freeze area Aspose.Cells | reusable worksheet freeze panes method | Aspose.Cells freeze panes example C#

using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    // The sample defines a reusable ApplyFreezePanes method that takes a Worksheet and zero‑based row/column indexes, invokes Worksheet.FreezePanes(row, column, row, column) to lock the specified area, fills sample data, freezes the first row and first two columns, and saves the workbook as FreezePanesResult.xlsx.
    class Program
    {
        // Reusable method that applies FreezePanes to the given worksheet.
        // row and column are 0‑based indexes where the freeze should start.
        static void ApplyFreezePanes(Worksheet sheet, int row, int column)
        {
            // Use the 4‑parameter overload: split at (row, column) and freeze the same number of rows/columns.
            sheet.FreezePanes(row, column, row, column);
        }

        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data to visualize the freeze effect.
                for (int i = 0; i < 20; i++)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                    }
                }

                // Freeze the first row and first two columns (row index 1, column index 2).
                ApplyFreezePanes(sheet, 1, 2);

                // Save the workbook.
                workbook.Save("FreezePanesResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
