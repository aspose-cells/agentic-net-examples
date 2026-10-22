// Title: Freeze rows and columns in an Excel file from command‑line arguments using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads two integer arguments from the command line, creates a workbook with Aspose.Cells, fills it with sample data, and freezes panes at the supplied row and column indices. | Show how to determine the worksheet's total row and column count and invoke Worksheet.FreezePanes(totalRows, totalColumns, freezeRow, freezeColumn) before saving the workbook as an .xlsx file.
// Common Searches: aspocells c# freeze panes based on command line input | how to use Worksheet.FreezePanes with dynamic row and column indices in .NET | C# program to freeze top rows and left columns in an Excel file using Aspose.Cells | command line arguments for FreezePanes method Aspose.Cells example | freeze panes at specific cell position Aspose.Cells C# tutorial
// Tags: Worksheet.FreezePanes with command‑line parameters | Aspose.Cells freeze rows and columns | populate worksheet before freezing panes | save workbook as .xlsx after FreezePanes | calculate total rows columns Aspose.Cells

using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    // The example reads two integers from the command line representing the row and column to freeze, creates a new workbook, populates cells with sample data, computes the worksheet's total rows and columns, applies Worksheet.FreezePanes using those totals and the user‑specified indices, saves the result as FrozenPaneOutput.xlsx, and prints a confirmation message.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate command‑line arguments: first is row index, second is column index.
                if (args.Length < 2 ||
                    !int.TryParse(args[0], out int freezeRow) ||
                    !int.TryParse(args[1], out int freezeColumn))
                {
                    Console.WriteLine("Usage: FreezePanesExample <freezeRow> <freezeColumn>");
                    return;
                }

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data to visualize the freeze effect.
                for (int r = 0; r < 20; r++)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                    }
                }

                // Apply FreezePanes.
                // The overload requires total rows/columns of the sheet plus the rows/columns to freeze.
                int totalRows = sheet.Cells.MaxDataRow + 1;
                int totalColumns = sheet.Cells.MaxDataColumn + 1;
                sheet.FreezePanes(totalRows, totalColumns, freezeRow, freezeColumn);

                // Save the workbook.
                string outputPath = "FrozenPaneOutput.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved with panes frozen at row {freezeRow}, column {freezeColumn}.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
